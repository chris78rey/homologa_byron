# Homologación ISSFA - Fase 01

Aplicación PyQt6 para homologar items en `SIS.ITEMS_ISSFA_DETALLE` y `SIS.EQUIVALENCIAS_ITEMS_ISSFA`.

## Requisitos

- Python 3.10+
- Oracle Instant Client + JDBC driver (`jdbc/ojdbc8.jar`)
- Variables de entorno configuradas

## Instalación

```bash
python3 -m venv venv
source venv/bin/activate
pip install PyQt6 jaydebeapi JPype1 openpyxl python-Levenshtein jellyfish
```

## Uso

```bash
source venv/bin/activate
export JAVA_TOOL_OPTIONS="-Doracle.jdbc.timezoneAsRegion=false -Duser.timezone=UTC"
export ORACLE_USER="tu_usuario"
export ORACLE_PASSWORD="tu_clave"
export ORACLE_TARGETS="172.16.60.20:1521:prdsgh1,172.16.60.21:1521:prdsgh2"
python main.py
```

## Flujo de Trabajo

1. **Login**: Ingresar credenciales Oracle (solo una vez)
2. **Plantilla**: Descargar plantilla oficial (📥)
3. **Excel**: Llenar plantilla y seleccionar archivo (📁)
4. **Configurar**: ID_ITISF y threshold de similitud
5. **Analizar**: Clic en "Analizar" para ver preview (🔍)
6. **Revisar**: Marcar/desmarcar filas según necesidad
7. **Aplicar**: Confirmar y aplicar cambios (✅)
8. **CSV**: Generar auditoría (📊)

## Formato Excel

| Columna | Descripción |
|---------|-------------|
| CODIGO_ACTUAL | Código actual en Oracle |
| DESCRIPCION_ACTUAL | Descripción actual (para comparar) |
| CODIGO_NUEVO | Nuevo código ISSFA |
| DESCRIPCION_NUEVA | Nueva descripción |

La plantilla oficial se descarga desde el botón **"Descargar plantilla Excel"**.

## Similitud (Jaro-Winkler + Levenshtein)

| Score | Decisión |
|-------|----------|
| 97-100% | Alta confianza (verde) |
| 88-96% | Revisar (amarillo) |
| <88% | No aplicar (rojo) |

## Vista Previa

| Acción | Descripción |
|--------|-------------|
| UPDATE | Actualizar código y descripción |
| INSERT | Insertar nuevo registro |
| BLOQUEADO | No se puede aplicar (ya existe o tiene equivalencias) |

## Tablas Involucradas

- **Lectura**: `SIS.ITEMS` (referencia)
- **Escritura**: `SIS.ITEMS_ISSFA_DETALLE`
- **Relaciones**: `SIS.EQUIVALENCIAS_ITEMS_ISSFA`
- **NO TOCAR**: `SIS.ITEMS_ISSFA_CABECERA`

## Backup y Restauración

Los backups se crean automáticamente:
- `SIS.BKP_ITEMS_ISSFA_DETALLE_YYYYMMDD`
- `SIS.BKP_EQUIVALENCIAS_ITEMS_ISSFA_YYYYMMDD`

## Estructura del Proyecto

```
homologa_byron/
├── main.py              # Interfaz PyQt6
├── database.py          # Conexión Oracle
├── homology.py          # Lógica de homologación
├── config.py            # Configuración
├── jdbc/                # Driver JDBC Oracle
├── resources/
│   └── templates/       # Plantilla oficial
│       └── plantilla_homologacion_items_issfa.xlsx
├── scripts/
│   └── create_template.py
├── crear_backup.sql     # Script SQL de backup
├── venv/                # Entorno virtual
└── README.md
```

## Reglas de Negocio

1. Si CODIGO_ACTUAL existe → UPDATE
2. Si CODIGO_ACTUAL no existe y CODIGO_NUEVO no existe → INSERT
3. Si CODIGO_NUEVO ya existe → BLOQUEADO
4. Si tiene equivalencias → BLOQUEADO (cambio de código peligroso)
5. Siempre vista previa antes de aplicar
6. Rollback automático si hay error
