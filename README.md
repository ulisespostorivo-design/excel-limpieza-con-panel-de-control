# Excel: Limpieza con Panel de Control (Auditoria y Control de Calidad de Datos)

## 📌 Descripcion del Proyecto

Este proyecto de Excel se enfoca en Data Quality (control de calidad de datos) y en la deteccion de inconsistencias antes de realizar la limpieza. El objetivo es evitar correcciones innecesarias y mantener un registro de los problemas encontrados.

Se desarrollo un panel de control que utiliza formulas para detectar, contar y clasificar diferentes tipos de errores. Los indicadores se actualizan cuando cambian los datos y permiten identificar que controles presentan problemas y cuantos registros estan afectados.

El archivo tambien incluye una guia con instrucciones para localizar los registros incorrectos, corregirlos y comprobar los resultados. El panel informa sobre los problemas detectados, pero no corrige automaticamente los datos.

---

## 📂 Estructura de los Archivos del Repositorio

El repositorio contiene tres archivos que muestran las distintas etapas del proceso:

1. **`01_raw_data.csv`**: conjunto de datos original, sin limpiar, utilizado como punto de partida.

2. **`02_audit_dashboard.xlsx`**: archivo de Excel con el panel de control, las formulas de validacion, las listas de valores permitidos y una guia para realizar la limpieza manualmente.

3. **`03_clean_data.csv`**: version final del conjunto de datos despues de aplicar las correcciones. Esta preparada para su importacion a una base de datos o para continuar con su analisis.

---

## 📊 Estructura del Dataset

El conjunto de datos simula registros de empleados e incluye los siguientes campos:

- `employee_id`: identificador unico del empleado.
- `first_name`: nombre.
- `last_name`: apellido.
- `gender`: genero.
- `age`: edad.
- `department`: departamento.
- `job_title`: cargo o puesto de trabajo.
- `salary`: salario.
- `hire_date`: fecha de contratacion.
- `email`: correo electronico.
- `city`: ciudad de residencia.
- `performance_rating`: calificacion de desempeño.

---

## 🛠️ El Panel de Control y las Formulas de Auditoria

El panel utiliza principalmente `SUMAPRODUCTO`, combinada con otras funciones de Excel, para contar los registros que incumplen las reglas definidas. Los indicadores del panel permiten identificar los tipos de problemas que requieren revision. Los valores superiores a cero se resaltan visualmente en el panel.

### 1. Deteccion de duplicados

Identifica las filas cuyo `employee_id` aparece mas de una vez. Los identificadores vacios no se cuentan.

```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*(CONTAR.SI(Datos[employee_id],Datos[employee_id])>1))
```

El contador incluye todas las apariciones de un identificador repetido, incluida la primera. No representa solamente la cantidad de copias adicionales.

### 2. Deteccion de espacios sobrantes

Compara el contenido original con el resultado de `ESPACIOS` para detectar espacios al principio, al final o repetidos entre palabras.

Ejemplo aplicado a `first_name`:

```excel
=SUMAPRODUCTO((Datos[first_name]<>"")*(Datos[first_name]<>ESPACIOS(Datos[first_name])))
```

El mismo control puede aplicarse por separado a otras columnas de texto.

### 3. Validacion de mayusculas

Compara los nombres con el formato generado por `NOMPROPIO` para detectar diferencias en el uso de mayusculas y minusculas.

Ejemplo aplicado a `first_name`:

```excel
=SUMAPRODUCTO((Datos[first_name]<>"")*(NO(IGUAL(Datos[first_name],NOMPROPIO(Datos[first_name])))))
```

Esta comprobacion detecta diferencias respecto al formato esperado, pero no determina por si sola si un nombre esta correctamente escrito.

### 4. Valores no permitidos

Compara los cargos con la lista de valores permitidos de `Permitidos[job_title]`. Detecta valores que no aparecen en esa lista, incluidos posibles errores de escritura y categorias que todavia no se hayan incorporado a ella.

```excel
=SUMAPRODUCTO((Datos[job_title]<>"")*(CONTAR.SI(Permitidos[job_title],Datos[job_title])=0))
```

Un valor detectado debe revisarse antes de corregirse. Puede tratarse de un error o de una categoria valida que falte en la lista.

### 5. Validacion de identificadores

Comprueba que los identificadores no vacios tengan seis caracteres, comiencen con la letra `E` y que los cinco caracteres restantes puedan convertirse en un numero.

```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*((LARGO(Datos[employee_id])<>6)+(IZQUIERDA(Datos[employee_id],1)<>"E")+ESERROR(VALOR(EXTRAE(Datos[employee_id],2,5)))>0))
```

La regla esperada en este dataset es una letra `E` seguida de cinco digitos. Cada identificador debe ser unico dentro del conjunto de datos. La ausencia de duplicados se comprueba mediante el control correspondiente.

### 6. Validacion basica del correo electronico

Comprueba que los correos no vacios contengan exactamente un simbolo `@` y que exista un punto despues de este.

```excel
=SUMAPRODUCTO((Datos[email]<>"")*((LARGO(Datos[email])-LARGO(SUSTITUIR(Datos[email],"@",""))<>1)+(SI.ERROR(HALLAR(".",Datos[email],HALLAR("@",Datos[email])+1),0)=0)>0))
```

Esta es una comprobacion basica de estructura. No garantiza que el correo exista ni que pueda recibir mensajes.

### 7. Validacion numerica de salarios

Detecta salarios no vacios almacenados como texto, valores numericos negativos y el valor especifico `999999`, definido como sospechoso para este dataset.

```excel
=SUMAPRODUCTO((Datos[salary]<>"")*((NO(ESNUMERO(Datos[salary]))+SI.ERROR(Datos[salary]<0,0)+SI.ERROR(Datos[salary]=999999,0))>0))
```

El control no detecta todos los salarios que podrian ser inusuales. El valor `999999` se comprueba porque fue definido expresamente como sospechoso. Los simbolos de moneda aplicados mediante el formato visual de Excel no convierten un numero en texto; los simbolos escritos dentro del propio valor pueden hacer que Excel lo interprete como texto.

### 8. Validacion de fechas de contratacion

Comprueba que las fechas no vacias sean reconocidas como valores numericos por Excel y que su representacion con el formato `dd/mm/yyyy` tenga diez caracteres.

```excel
=SUMAPRODUCTO((Datos[hire_date]<>"")*((NO(ESNUMERO(Datos[hire_date]))+(LARGO(TEXTO(Datos[hire_date],"dd/mm/yyyy"))<>10))>0))
```

Este control permite detectar valores que Excel no reconoce como numeros de fecha. No comprueba por si solo si una fecha es coherente con la historia laboral del empleado.

### 9. Validacion numerica de edades

Detecta edades no vacias que no estan almacenadas como numeros.

```excel
=SUMAPRODUCTO((Datos[age]<>"")*(NO(ESNUMERO(Datos[age])))
```

El control revisa el tipo de dato, pero no determina si una edad numerica es razonable. Esa evaluacion requiere una regla adicional.

---

## 🎯 Objetivo del Proyecto

Demostrar el uso de Excel para auditar la calidad de un conjunto de datos, detectar y cuantificar inconsistencias, documentar un procedimiento de limpieza y obtener una version corregida que pueda utilizarse en etapas posteriores de analisis.

El proyecto busca que el proceso sea comprensible, verificable y adaptable a otros conjuntos de datos mediante la definicion de reglas especificas para cada columna.



