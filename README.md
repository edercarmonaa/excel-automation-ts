# excel-automation-ts

📊 Repositorio de scripts para automatizar tareas en Excel usando TypeScript y Office Scripts. Incluye ejemplos y soluciones para optimizar flujos de trabajo en hojas de cálculo. 🚀

Scripts desarrollados en TypeScript para su uso con Office Scripts en Excel Online. Contiene ejemplos prácticos de manipulación de datos, automatización de tareas y personalización de hojas de cálculo.

##Indice

- [account_number](account_number.osts)
- [bank_support](bank_support.osts)
- [charge_credit](charge_credit.osts)
- [combine](combine.osts)
- [copy_paste](copy_paste.osts)
- [date](date.osts)
- [del_col](del_col.osts)

## copy_paste

Este script permite copiar el contenido de un rango seleccionado en la hoja activa y pegarlo en todas las demás hojas del libro en la misma posición. Es útil cuando necesitas replicar un conjunto de datos o formatos de manera rápida y consistente en múltiples hojas dentro del mismo libro de Excel.

## del_col
Este código es útil cuando se necesita limpiar el contenido de un archivo de Excel eliminando las columnas C a L en todas las hojas. Ideal para procesar datos y reducir información innecesaria antes de análisis o exportación.

## bank_support
Este script toma un rango de celdas seleccionado por el usuario, y con base en el banco y la fecha, asigna un valor específico a dos columnas del rango. El valor asignado sigue un patrón específico que depende del banco (por ejemplo, "SANTANDER" genera un valor como "CTA_2787_mm" donde "mm" es el mes de la fecha proporcionada). Este proceso se realiza para todas las filas del rango seleccionado.

## date
La funcionalidad de este script de Excel es convertir las fechas en un rango seleccionado de la hoja activa, de un formato específico (que parece ser 'D/MM/YYYY', con un día de dos dígitos seguido de un mes de dos dígitos y un año de cuatro dígitos) a un nuevo formato ('MM/DD/YYYY').

## combine
Este script de ExcelScript está diseñado para combinar las descripciones de filas consecutivas con celdas vacías en la columna de fechas. El contenido de las filas vacías se combina en la fila de la fecha vacía hasta que se encuentre una fila con una fecha válida.

## charge_credit
Este script en ExcelScript realiza varias operaciones sobre el rango seleccionado: elimina los apóstrofes de las celdas, aplica un formato numérico con separación de miles y dos decimales, y ajusta los valores de las columnas moviendo el contenido de la segunda columna a la primera si no es un + o dejando la celda vacía si es un +. Finalmente, reasigna los valores modificados y garantiza que el formato numérico se mantenga consistente con dos decimales.

## account_number
Este script en ExcelScript asigna un número de cuenta específico a las filas del rango seleccionado, según el banco indicado en una columna. Si el banco no está vacío, el script verifica el nombre del banco y asigna el número de cuenta correspondiente a la columna de cuenta. Los bancos incluidos en el script son "SERFIN", "BBVA", "CITI", "HSBC", "BANORTE" y "COMPE". Finalmente, los valores actualizados se asignan nuevamente al rango.

