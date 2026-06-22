-- =========================================================
-- CREAR Y SELECCIONAR BASE DE DATOS
-- MYSQL 8 / MARIADB / XAMPP / PHPMYADMIN
-- =========================================================

CREATE DATABASE IF NOT EXISTS mantenimiento_analytics;

USE mantenimiento_analytics;


-- =========================================================
-- ELIMINAR TABLAS SI EXISTEN
-- =========================================================

DROP TABLE IF EXISTS fact_ordenes_trabajo;
DROP TABLE IF EXISTS dim_tecnicos;
DROP TABLE IF EXISTS dim_equipos;
DROP TABLE IF EXISTS dim_tipo_falla;
DROP TABLE IF EXISTS dim_prioridad;


-- =========================================================
-- TABLAS DE DIMENSIÓN
-- =========================================================

CREATE TABLE dim_tecnicos (

    id_tecnico VARCHAR(10) PRIMARY KEY,

    nombre VARCHAR(100) NOT NULL,

    especialidad VARCHAR(50) NOT NULL,

    anios_experiencia INT NOT NULL
        CHECK (anios_experiencia >= 0),

    costo_hora DECIMAL(10,2) NOT NULL
        CHECK (costo_hora > 0)

);


CREATE TABLE dim_equipos (

    id_equipo INT AUTO_INCREMENT PRIMARY KEY,

    codigo VARCHAR(20) UNIQUE NOT NULL,

    descripcion VARCHAR(100),

    tipo VARCHAR(50),

    ubicacion VARCHAR(100),

    criticidad VARCHAR(10)
        CHECK (
            criticidad IN ('Alta','Media','Baja')
        )

);


CREATE TABLE dim_tipo_falla (

    id_tipo_falla INT AUTO_INCREMENT PRIMARY KEY,

    nombre VARCHAR(50) UNIQUE NOT NULL,

    categoria VARCHAR(20)
        CHECK (
            categoria IN (
                'Planificada',
                'No Planificada'
            )
        )

);


CREATE TABLE dim_prioridad (

    id_prioridad INT AUTO_INCREMENT PRIMARY KEY,

    nivel VARCHAR(10) UNIQUE NOT NULL,

    orden_sla INT

);


-- =========================================================
-- TABLA DE HECHOS
-- =========================================================

CREATE TABLE fact_ordenes_trabajo (

    id_orden VARCHAR(10) PRIMARY KEY,

    fecha_creacion DATE NOT NULL,

    fecha_cierre DATE,

    id_tipo_falla INT,

    id_equipo INT,

    id_tecnico VARCHAR(10),

    id_prioridad INT,

    horas_trabajadas DECIMAL(8,2)
        CHECK (horas_trabajadas >= 0),

    costo_mano_obra DECIMAL(12,2) DEFAULT 0,

    costo_repuestos DECIMAL(12,2) DEFAULT 0,

    costo_logistica DECIMAL(12,2) DEFAULT 0,

    costo_terceros DECIMAL(12,2) DEFAULT 0,

    costo_parada DECIMAL(12,2) DEFAULT 0,

    costo_penalizacion DECIMAL(12,2) DEFAULT 0,

    costo_total DECIMAL(14,2)
        GENERATED ALWAYS AS (

            COALESCE(costo_mano_obra,0)
          + COALESCE(costo_repuestos,0)
          + COALESCE(costo_logistica,0)
          + COALESCE(costo_terceros,0)
          + COALESCE(costo_parada,0)
          + COALESCE(costo_penalizacion,0)

        ) STORED,

    estado VARCHAR(15)
        CHECK (
            estado IN (
                'Cerrada',
                'En Proceso',
                'Pendiente'
            )
        ),

    tiempo_resolucion_horas DECIMAL(8,1),

    retrabajo BOOLEAN DEFAULT FALSE,

    observaciones TEXT,

    CONSTRAINT chk_fechas
        CHECK (
            fecha_cierre IS NULL
            OR fecha_cierre >= fecha_creacion
        ),

    CONSTRAINT fk_tipo_falla
        FOREIGN KEY (id_tipo_falla)
        REFERENCES dim_tipo_falla(id_tipo_falla),

    CONSTRAINT fk_equipo
        FOREIGN KEY (id_equipo)
        REFERENCES dim_equipos(id_equipo),

    CONSTRAINT fk_tecnico
        FOREIGN KEY (id_tecnico)
        REFERENCES dim_tecnicos(id_tecnico),

    CONSTRAINT fk_prioridad
        FOREIGN KEY (id_prioridad)
        REFERENCES dim_prioridad(id_prioridad)

);


-- =========================================================
-- ÍNDICES PARA PERFORMANCE ANALÍTICA
-- =========================================================

CREATE INDEX idx_ot_fecha
ON fact_ordenes_trabajo(fecha_creacion);

CREATE INDEX idx_ot_tecnico
ON fact_ordenes_trabajo(id_tecnico);

CREATE INDEX idx_ot_equipo
ON fact_ordenes_trabajo(id_equipo);

CREATE INDEX idx_ot_tipo
ON fact_ordenes_trabajo(id_tipo_falla);

CREATE INDEX idx_ot_estado
ON fact_ordenes_trabajo(estado);

CREATE INDEX idx_ot_prioridad
ON fact_ordenes_trabajo(id_prioridad);

CREATE INDEX idx_ot_retrabajo
ON fact_ordenes_trabajo(retrabajo);


-- =========================================================
-- DATOS MAESTROS
-- =========================================================

INSERT INTO dim_tipo_falla (nombre, categoria)
VALUES
    ('Eléctrica',   'No Planificada'),
    ('Mecánica',    'No Planificada'),
    ('Preventiva',  'Planificada'),
    ('Correctiva',  'No Planificada'),
    ('Emergencia',  'No Planificada');


INSERT INTO dim_prioridad (nivel, orden_sla)
VALUES
    ('Alta',  1),
    ('Media', 2),
    ('Baja',  3);


-- =========================================================
-- Q1: MTTR POR TIPO DE FALLA
-- KPI FUNDAMENTAL DE MANTENIMIENTO
-- =========================================================

SELECT

    tf.nombre AS tipo_falla,

    COUNT(*) AS total_ordenes,

    ROUND(
        AVG(ot.tiempo_resolucion_horas),
        1
    ) AS mttr_horas,

    ROUND(
        STDDEV(ot.tiempo_resolucion_horas),
        1
    ) AS desv_estandar,

    ROUND(
        MIN(ot.tiempo_resolucion_horas),
        1
    ) AS min_horas,

    ROUND(
        MAX(ot.tiempo_resolucion_horas),
        1
    ) AS max_horas

FROM fact_ordenes_trabajo ot

JOIN dim_tipo_falla tf
ON ot.id_tipo_falla = tf.id_tipo_falla

WHERE ot.estado = 'Cerrada'
AND ot.tiempo_resolucion_horas IS NOT NULL

GROUP BY tf.nombre

ORDER BY mttr_horas DESC;


-- =========================================================
-- Q2: RANKING DE EFICIENCIA POR TÉCNICO
-- SCORECARD EJECUTIVO
-- =========================================================

WITH base AS (

    SELECT

        t.id_tecnico,

        t.nombre,

        t.especialidad,

        t.anios_experiencia,

        COUNT(ot.id_orden) AS total_ot,

        AVG(ot.horas_trabajadas) AS horas_prom,

        AVG(ot.costo_total) AS costo_prom,

        AVG(
            CASE
                WHEN ot.retrabajo = TRUE THEN 1
                ELSE 0
            END
        ) AS ratio_retrabajo,

        AVG(ot.tiempo_resolucion_horas) AS mttr_prom

    FROM fact_ordenes_trabajo ot

    JOIN dim_tecnicos t
    ON ot.id_tecnico = t.id_tecnico

    WHERE ot.estado = 'Cerrada'

    GROUP BY
        t.id_tecnico,
        t.nombre,
        t.especialidad,
        t.anios_experiencia
),

normalizado AS (

    SELECT

        *,

        MAX(horas_prom) OVER() AS max_horas,

        MAX(costo_prom) OVER() AS max_costo,

        MAX(ratio_retrabajo) OVER() AS max_retrabajo

    FROM base
)

SELECT

    nombre AS tecnico,

    especialidad,

    anios_experiencia AS exp_anios,

    total_ot,

    ROUND(horas_prom,1) AS horas_prom,

    ROUND(costo_prom,0) AS costo_prom_clp,

    ROUND(ratio_retrabajo * 100,1) AS pct_retrabajo,

    ROUND(mttr_prom,1) AS mttr_prom,

    ROUND(

        (
            (horas_prom / NULLIF(max_horas,0)) * 0.35
          + (costo_prom / NULLIF(max_costo,0)) * 0.40
          + (ratio_retrabajo / NULLIF(max_retrabajo,0)) * 0.25
        ) * 100

    ,1) AS ineficiencia_score

FROM normalizado

ORDER BY ineficiencia_score ASC;


-- =========================================================
-- Q3: PARETO DE EQUIPOS
-- =========================================================

WITH equipo_stats AS (

    SELECT

        ot.id_equipo,

        COUNT(*) AS frecuencia,

        SUM(ot.costo_total) AS costo_acumulado,

        AVG(ot.costo_total) AS costo_promedio,

        SUM(
            CASE
                WHEN ot.retrabajo = TRUE THEN 1
                ELSE 0
            END
        ) AS total_retrabajos

    FROM fact_ordenes_trabajo ot

    GROUP BY ot.id_equipo
),

totales AS (

    SELECT
        SUM(costo_acumulado) AS gran_total

    FROM equipo_stats
)

SELECT

    e.codigo AS equipo,

    e.ubicacion,

    e.criticidad,

    es.frecuencia,

    ROUND(es.costo_acumulado,0) AS costo_total_clp,

    ROUND(es.costo_promedio,0) AS costo_prom_clp,

    ROUND(
        (
            es.costo_acumulado
            / NULLIF(t.gran_total,0)
        ) * 100
    ,1) AS pct_costo_total,

    ROUND(

        (
            SUM(es.costo_acumulado)
            OVER (
                ORDER BY es.costo_acumulado DESC
            )

            / NULLIF(t.gran_total,0)
        ) * 100

    ,1) AS pct_acumulado,

    es.total_retrabajos

FROM equipo_stats es

JOIN dim_equipos e
ON es.id_equipo = e.id_equipo

CROSS JOIN totales t

ORDER BY es.costo_acumulado DESC;


-- =========================================================
-- Q4: TENDENCIA MENSUAL DE COSTOS Y VOLUMEN
-- =========================================================

SELECT

    DATE_FORMAT(
        ot.fecha_creacion,
        '%Y-%m-01'
    ) AS mes,

    COUNT(*) AS ordenes_creadas,

    SUM(
        CASE
            WHEN ot.estado = 'Cerrada'
            THEN 1
            ELSE 0
        END
    ) AS ordenes_cerradas,

    ROUND(
        SUM(ot.costo_total),
        0
    ) AS costo_total_mes,

    ROUND(
        AVG(ot.costo_total),
        0
    ) AS costo_promedio,

    SUM(
        CASE
            WHEN ot.retrabajo = TRUE
            THEN 1
            ELSE 0
        END
    ) AS retrabajos_mes,

    ROUND(
        AVG(ot.horas_trabajadas),
        1
    ) AS horas_prom

FROM fact_ordenes_trabajo ot

GROUP BY
    DATE_FORMAT(
        ot.fecha_creacion,
        '%Y-%m'
    )

ORDER BY mes;


-- =========================================================
-- Q5: DETECCIÓN DE OUTLIERS
-- COSTOS > MEDIA + 2.5 DESV ESTÁNDAR
-- =========================================================

WITH estadisticas AS (

    SELECT

        AVG(costo_total) AS media,

        STDDEV(costo_total) AS desv

    FROM fact_ordenes_trabajo

    WHERE estado = 'Cerrada'
)

SELECT

    ot.id_orden,

    ot.fecha_creacion,

    tf.nombre AS tipo_falla,

    t.nombre AS tecnico,

    ot.horas_trabajadas,

    ROUND(ot.costo_total,0) AS costo_total_clp,

    ROUND(
        (
            ot.costo_total - e.media
        )
        /
        NULLIF(e.desv,0)
    ,2) AS z_score,

    ot.retrabajo

FROM fact_ordenes_trabajo ot

JOIN estadisticas e
ON TRUE

JOIN dim_tipo_falla tf
ON ot.id_tipo_falla = tf.id_tipo_falla

JOIN dim_tecnicos t
ON ot.id_tecnico = t.id_tecnico

WHERE ot.estado = 'Cerrada'

AND ot.costo_total >

(
    COALESCE(e.media,0)
    +
    (
        2.5 * COALESCE(e.desv,0)
    )
)

ORDER BY ot.costo_total DESC

LIMIT 50;


-- =========================================================
-- Q6: % CUMPLIMIENTO SLA POR PRIORIDAD
-- =========================================================

SELECT

    p.nivel AS prioridad,

    COUNT(*) AS total,

    SUM(

        CASE

            WHEN p.nivel = 'Alta'
            AND ot.tiempo_resolucion_horas <= 24
            THEN 1

            WHEN p.nivel = 'Media'
            AND ot.tiempo_resolucion_horas <= 72
            THEN 1

            WHEN p.nivel = 'Baja'
            AND ot.tiempo_resolucion_horas <= 168
            THEN 1

            ELSE 0

        END

    ) AS cumple_sla,

    ROUND(

        (
            SUM(

                CASE

                    WHEN p.nivel = 'Alta'
                    AND ot.tiempo_resolucion_horas <= 24
                    THEN 1

                    WHEN p.nivel = 'Media'
                    AND ot.tiempo_resolucion_horas <= 72
                    THEN 1

                    WHEN p.nivel = 'Baja'
                    AND ot.tiempo_resolucion_horas <= 168
                    THEN 1

                    ELSE 0

                END

            )

            / NULLIF(COUNT(*),0)

        ) * 100

    ,1) AS pct_sla

FROM fact_ordenes_trabajo ot

JOIN dim_prioridad p
ON ot.id_prioridad = p.id_prioridad

WHERE ot.estado = 'Cerrada'
AND ot.tiempo_resolucion_horas IS NOT NULL

GROUP BY
    p.nivel,
    p.orden_slaid_prioridadfact_ordenes_trabajodim_prioridadid_tecnicoid_tecnico

ORDER BY p.orden_sla;codigo