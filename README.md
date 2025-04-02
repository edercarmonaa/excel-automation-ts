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
- [new_sheet](new_sheet.osts)
- [random](random.osts)
- [school_grades](school_grades.osts)
- [width](width.osts)

## account_number
Este script en ExcelScript asigna un número de cuenta específico a las filas del rango seleccionado, según el banco indicado en una columna. Si el banco no está vacío, el script verifica el nombre del banco y asigna el número de cuenta correspondiente a la columna de cuenta. Los bancos incluidos en el script son "SERFIN", "BBVA", "CITI", "HSBC", "BANORTE" y "COMPE". Finalmente, los valores actualizados se asignan nuevamente al rango.

## bank_support
Este script toma un rango de celdas seleccionado por el usuario, y con base en el banco y la fecha, asigna un valor específico a dos columnas del rango. El valor asignado sigue un patrón específico que depende del banco (por ejemplo, "SANTANDER" genera un valor como "CTA_2787_mm" donde "mm" es el mes de la fecha proporcionada). Este proceso se realiza para todas las filas del rango seleccionado.

## charge_credit
Este script en ExcelScript realiza varias operaciones sobre el rango seleccionado: elimina los apóstrofes de las celdas, aplica un formato numérico con separación de miles y dos decimales, y ajusta los valores de las columnas moviendo el contenido de la segunda columna a la primera si no es un + o dejando la celda vacía si es un +. Finalmente, reasigna los valores modificados y garantiza que el formato numérico se mantenga consistente con dos decimales.

## combine
Este script de ExcelScript está diseñado para combinar las descripciones de filas consecutivas con celdas vacías en la columna de fechas. El contenido de las filas vacías se combina en la fila de la fecha vacía hasta que se encuentre una fila con una fecha válida.

## copy_paste
Este script permite copiar el contenido de un rango seleccionado en la hoja activa y pegarlo en todas las demás hojas del libro en la misma posición. Es útil cuando necesitas replicar un conjunto de datos o formatos de manera rápida y consistente en múltiples hojas dentro del mismo libro de Excel.

## date
La funcionalidad de este script de Excel es convertir las fechas en un rango seleccionado de la hoja activa, de un formato específico (que parece ser 'D/MM/YYYY', con un día de dos dígitos seguido de un mes de dos dígitos y un año de cuatro dígitos) a un nuevo formato ('MM/DD/YYYY').

## del_col
Este código es útil cuando se necesita limpiar el contenido de un archivo de Excel eliminando las columnas C a L en todas las hojas. Ideal para procesar datos y reducir información innecesaria antes de análisis o exportación.

##new_sheet
Este script en ExcelScript recorre los valores dentro del rango seleccionado por el usuario en la hoja activa. Para cada valor (que debe ser un nombre de hoja válido y no vacío), verifica si ya existe una hoja con ese nombre. Si la hoja no existe, crea una nueva hoja con ese nombre. Si ya existe, muestra un mensaje en la consola indicando que la hoja ya está presente.

## random
Este script en ExcelScript selecciona un número aleatorio de celdas dentro de un rango especificado en la hoja activa. El número de celdas a seleccionar es pasado como argumento n. Primero, verifica que n no exceda el número total de celdas en el rango. Luego, selecciona n celdas aleatorias y les aplica un formato (en este caso, un fondo amarillo). Para lograrlo, el script calcula índices aleatorios, los convierte en coordenadas de fila y columna, y aplica el formato deseado a esas celdas seleccionadas.

## school_grades
Este script en ExcelScript toma un rango de calificaciones numéricas seleccionadas en una hoja de cálculo y convierte cada calificación en su equivalente en palabras en español. Primero, obtiene las calificaciones numéricas del rango seleccionado y las convierte a palabras usando la función numberToWords. Esta función maneja números del 0 al 99, cubriendo unidades, decenas y números entre 10 y 19. Luego, almacena las calificaciones en palabras en un rango específico (de la celda B1 a B10). Así, el script convierte las calificaciones numéricas a su representación textual y las coloca en un nuevo rango en la hoja de cálculo.

## width
Este script de ExcelScript recorre todas las hojas de un libro de Excel y ajusta el formato de las columnas B, C, D y E. Primero, obtiene todas las hojas del libro y luego iterativamente accede a las columnas especificadas. Para cada columna, se desactiva el ajuste de texto (wrap text) y se autoajusta el ancho de la columna para que se adapte al contenido. El rango de cada columna se determina según las filas utilizadas en cada hoja. Al finalizar, se imprime un mensaje de confirmación en la consola.
