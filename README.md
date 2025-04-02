# excel-automation-ts

📊 Repositorio de scripts para automatizar tareas en Excel usando TypeScript y Office Scripts. Incluye ejemplos y soluciones para optimizar flujos de trabajo en hojas de cálculo. 🚀

Scripts desarrollados en TypeScript para su uso con Office Scripts en Excel Online. Contiene ejemplos prácticos de manipulación de datos, automatización de tareas y personalización de hojas de cálculo.

## copy_paste

Este script permite copiar el contenido de un rango seleccionado en la hoja activa y pegarlo en todas las demás hojas del libro en la misma posición. Es útil cuando necesitas replicar un conjunto de datos o formatos de manera rápida y consistente en múltiples hojas dentro del mismo libro de Excel.

## del_col
Este código es útil cuando se necesita limpiar el contenido de un archivo de Excel eliminando las columnas C a L en todas las hojas. Ideal para procesar datos y reducir información innecesaria antes de análisis o exportación.

## bank_support
Este script toma un rango de celdas seleccionado por el usuario, y con base en el banco y la fecha, asigna un valor específico a dos columnas del rango. El valor asignado sigue un patrón específico que depende del banco (por ejemplo, "SANTANDER" genera un valor como "CTA_2787_mm" donde "mm" es el mes de la fecha proporcionada). Este proceso se realiza para todas las filas del rango seleccionado.

##date
La funcionalidad de este script de Excel es convertir las fechas en un rango seleccionado de la hoja activa, de un formato específico (que parece ser 'D/MM/YYYY', con un día de dos dígitos seguido de un mes de dos dígitos y un año de cuatro dígitos) a un nuevo formato ('MM/DD/YYYY').

##Indice
- [copy_paste](copy_paste.osts)
- [del_col](del_col.osts)
- [bank_support](bank_support.osts)
- [date](date.osts)
