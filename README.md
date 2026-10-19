# Excel: Limpieza con Panel de Control (Auditoría y Control de Calidad de Datos)

## 📌 Descripción del Proyecto

Este proyecto de Excel está enfocado en **Data Quality** (Control de Calidad de Datos) y resolución de problemas analíticos con un enfoque eficiente: **evitar las limpiezas a ciegas o preventivas**, las cuales consumen tiempo y recursos innecesarios.

En lugar de corregir datos sin rumbo, se diseñó un **Panel de Control de Auditoría** interactivo impulsado por fórmulas lógicas y matriciales. Este panel detecta, cuenta y categoriza automáticamente anomalías en tiempo real (con alertas visuales), permitiendo saber exactamente *dónde* está el error, *cuántos* hay y de *qué tipo*, para realizar una **limpieza guiada, rápida y precisa**.

---

## 📂 Estructura de los Archivos del Repositorio

El repositorio está organizado en tres etapas clave para mostrar la evolución y el proceso completo de auditoría:

1. **`01_raw_data.csv`**: El conjunto de datos original y sin procesar (crudo).


2. **`02_audit_dashboard.xlsx`**: El archivo que contiene el panel de control de auditoría, las fórmulas de validación, las tablas auxiliares (listas blancas/whitelists) y los registros con errores intencionales.


3. **`03_clean_data.csv`**: El dataset totalmente depurado, validado y listo para producción.



---

## 📊 Estructura del Dataset

El conjunto de datos simula registros de empleados corporativos e incluye los siguientes campos:

* `employee_id`: Identificador único del empleado.


* `first_name` / `last_name`: Nombre y apellido del trabajador.


* `gender`: Género del empleado (estandarizado).


* `age`: Edad.


* `department`: Departamento corporativo.


* `job_title`: Cargo o puesto de trabajo.


* `salary`: Salario asignado.


* `hire_date`: Fecha de contratación.


* `email`: Correo electrónico corporativo.


* `city`: Ciudad de residencia.


* `performance_rating`: Calificación de desempeño.



---

## 🛠️ El Panel de Control y Fórmulas de Auditoría

El núcleo de la herramienta utiliza la función `SUMAPRODUCTO` combinada con lógica condicional para generar contadores automáticos de errores por columna. Si el panel general detecta un valor superior a `0`, activa alertas visuales (en rojo) para advertir al operador de forma inmediata.

### 1. Detección de Duplicados

Identifica registros repetidos evitando falsos positivos en celdas vacías:

```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*(CONTAR.SI(Datos[employee_id],Datos[employee_id])>1))

```

### 2. Detección de Espacios Sobrantes

Compara el texto original con la función `ESPACIOS` para detectar espacios iniciales, finales o dobles internos no deseados:

```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*(Datos[employee_id]<>ESPACIOS(Datos[employee_id])))

```

### 3. Detección de Mayúsculas Incorrectas

Verifica que los nombres y apellidos sigan el formato de nombre propio adecuado:

```excel
=SUMAPRODUCTO((Datos[first_name]<>"")*(NO(IGUAL(Datos[first_name],NOMPROPIO(Datos[first_name])))))

```

### 4. Detección de Errores de Escritura (Listas Blancas / Whitelists)

Compara los valores del campo con una tabla auxiliar de puestos permitidos (`Permitidos[job_title]`) para encontrar entradas mal escritas o valores no autorizados de forma escalable:

```excel
=SUMAPRODUCTO((Datos[job_title]<>"")*(CONTAR.SI(Permitidos[job_title],Datos[job_title])=0))

```

### 5. Validación de Formatos y Tipos de Datos

* **Identificador de Empleado**: Valida que tenga exactamente 6 caracteres, comience con la letra "E" y le sigan 5 números:


```excel
=SUMAPRODUCTO((Datos[employee_id]<>"")*((LARGO(Datos[employee_id])<>6)+(IZQUIERDA(Datos[employee_id],1)<>"E")+ESERROR(VALOR(EXTRAE(Datos[employee_id],2,5)))>0))

```


* **Correo Electrónico**: Comprueba la estructura correcta verificando la presencia de un único símbolo `@` y un punto `.` posterior:


```excel
=SUMAPRODUCTO((Datos[email]<>"")*((LARGO(Datos[email])-LARGO(SUSTITUIR(Datos[email],"@",""))<>1)+(SI.ERROR(HALLAR(".",Datos[email],HALLAR("@",Datos[email])+1),0)=0)>0))

```


* **Salarios**: Detecta valores con formatos incorrectos, números negativos, símbolos erróneos o valores atípicos (como `999999`):


```excel
=SUMAPRODUCTO((Datos[salary]<>"")*((NO(ESNUMERO(Datos[salary]))+SI.ERROR(Datos[salary]<0,0)+SI.ERROR(Datos[salary]=999999,0))>0))

```


* **Fechas de Contratación**: Valida que el formato corresponda a una estructura numérica y de texto válida de 10 caracteres:


```excel
=SUMAPRODUCTO((Datos[hire_date]<>"")*((NO(ESNUMERO(Datos[hire_date]))+(LARGO(TEXTO(Datos[hire_date],"dd/mm/yyyy"))<>10))>0))

```



---

## 📝 Bitácora de Desarrollo

El proceso detallado de descubrimiento de errores, diseño del panel de control, resolución de conflictos con formatos internacionales (como fechas o géneros) y evolución de las fórmulas se encuentra documentado paso a paso en el registro del proyecto.
