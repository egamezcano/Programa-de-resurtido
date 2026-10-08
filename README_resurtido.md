# Automatización de resurtido

## Instalación

Se recomienda Python 3.11 o superior.

```bash
python -m venv .venv
```

Windows:

```bash
.venv\\Scripts\\activate
```

macOS/Linux:

```bash
source .venv/bin/activate
```

Instala dependencias:

```bash
pip install -r requirements.txt
```

## Ejecución

Coloca `inventario\_ferreteria\_garza.xlsx` en la carpeta del script y ejecuta:

```bash
python resurtido.py
```

Opciones:

```bash
python resurtido.py --input "ruta/inventario.xlsx" --orders "pedidos\_proveedores.xlsx" --exceptions "reporte\_excepciones.xlsx" --drafts "borradores\_correos" --log "logs/resurtido.log"
```

## Estructura esperada

El libro debe incluir una hoja de proveedores con columnas equivalentes a:

* `Proveedor`
* `Correo`
* `Pedido\_minimo\_MXN`

Opcionalmente: `Nombre\_proveedor` y `Contacto`.

Las otras hojas de datos de producto deben contener columnas equivalentes a:

* `Codigo`
* `Descripcion`
* `Existencia`
* `Minimo`
* `Maximo`
* `Activo`
* `Empaque\_multiplo`
* `Proveedor`
* `Costo\_unitario\_MXN`

Los nombres de columnas aceptan algunos alias en español; consulte `ALIASES` en `resurtido.py` para ampliarlos. Las claves de proveedor de las hojas de productos deben coincidir con la clave `Proveedor` de la hoja de proveedores.

## Salidas

* `pedidos\_proveedores.xlsx`: detalle y resumen por proveedor, incluido estado del mínimo.
* `reporte\_excepciones.xlsx`: filas/hojas omitidas, motivos y acciones.
* `borradores\_correos/\*.eml`: borradores listos para abrir y revisar.
* `logs/resurtido.log`: ejecución y pedidos retenidos por mínimo.

## Seguridad operativa

El programa **no envía correos**. Genera borradores `.eml` para revisión humana. Los pedidos marcados `RETENIDO\_POR\_MINIMO` no deben enviarse hasta que Compras decida consolidar más productos, negociar una excepción o autorizar otro tratamiento.

Antes de producción, valida los nombres de las hojas y columnas con una copia del libro real y realiza pruebas con datos representativos.

