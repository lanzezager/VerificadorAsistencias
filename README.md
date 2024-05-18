# Verificador de Asistencias

Esta es una aplicación que pretende facilitar la verificación de asistencia de los alumnos a los cursos de Zoom o a la plataforma, 
buscando para ello a los asistentes utilizando los siguientes elementos y en orden de prioridad:

1. Nombre y Apellidos Completos
2. Correo Electronico
3. Primer Nombre y Apellidos Completos
4. Segundo Nombre y Apellidos Completos
5. Primer Nombre y Primer Apellido
6. Segundo Nombre y Primer Apellido

## Instrucciones de Uso

![ListaFormatoSeguimiento](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEhqCVCKmonr3IrMWOp_H6FKyfDFEYskmQAqL0Rm3_2_LNauFRWloFpzHkYhKeVBmB75Lp1EvAgOxGfFBzqhRH3KY1C-WT63LG3DF-AGZBn989RFvkENTMclQvH82YVEGb8xk1TX0W8VRyIrC6DgKGiYaPiql9Dor-EL9gkHG-CVXZJyBnRkmXKZSWOybuI3/s967/Screenshot_1.png)

Para que la Aplicacion cumpla su objetivo necesita que se ingrese información a las tablas de datos, para ello se pueden utilizar la opción de click derecho **Pegar** o el comando de windows CTRL+V sobre el área del primer campo para añadir el listado de alumnos previamente copiado del formato de seguimiento, se deben copiar las columnas de **Nombre**, **Apellido** y **Correo Electronico** omitiendo el Encabezado para el correcto funcionamiento del software.

![ListaZoom-Plataforma](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEhMVuZb3epGN0vyVlpFNh-SqtkY8oW12LfOsrGQTknFhIlmsZp40IcYorlUJYwVRn6SbsJKQ3ln8NneJRiYxP3XXK-_lMqD2PFaOlHSNUIOLGLhaAoHc57NT4rpIc3jbHzgE55sZF4dF912LpdklGwqxbwJMn-e8Y_1Qhaw64rn2joQIWp8dxejVNCMcvVh/s967/Screenshot_2.png)

La lista de asistencia de Zoom se debe de añadir de la misma forma en el segundo campo y tomado las columnas de **Nombre Completo/Usuario**, y **Correo Electronico**, para la asistencia a la plataforma se deberán seleccionar las columna de **Usuario** y otra indiferente ya que se evaluará solo la coincidencia total del nombre de usuario registrado y utilzado para acceder a la plataforma.

![Listados](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEgldv0iZm11eotNZiV0uZhKUEIaxBBbTI0M-Xvuq_bRcAxEKx-eXsJvFlqOfG3JssmLEVE70VYj915epgjyjpcq8uuKhTQTYXUT34tC6s8nSK87JGBAB3pRyn7D7meegCMQmCgF7zoaySpJIaPjdqJf4zyf69bNKz2mB8zRIOhRxPrObx6oSOTotdvsDP3-/s970/Screenshot_3.png)

Una vez ingresadas ambas listas se deberá dar click en el botón **Verificar** para proceder con la revision de las asistencias.

![Resultados](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEg83lypH9oNZbt5V-or67cvgYi9PetCniBvNEm4yIhXVGzuNFQEwLtck2dt_k7JjOnUY4ksNyL1eewersS3-KYrR6Titkc6nwgHzP2keJZnO2F6wit6dwGqybVR-VCK0CChc2QXsY6PHonJoPZ5kuLUIavYa9P2QZ3jpiUjm-jSmiONZnBYKpyrfTKTNGHC/s968/Screenshot_4.png)

Cuando el proceso haya terminado aparecerá una ventana de dialogo informando el numero total de Asistencias encontradas.

![Evaluacion](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEhMkE294Q73oqxSj4qkFaqKrpXaW2QpxGSA8nJPxJjKEG2jSbflEpABKMK5MWr4Mbk_WI0jDvksVJhnuiGHidUFZKnWGXt8fOQxoPb5hekJGBJPbVqaoy7Ya4KKRhG_mvvRmF6d4Ol2loLh9_2yp5ZANGKPUB4zkKX_RF7wqTh7EiEVp3MfAhsI_JOVRAP-/s967/Screenshot_5.png)

En el listado de Resultados podrá identificar a los participantes que asistieron a la clase / curso en plataforma ya que el programa les habrá asignado un numero en la columna de ASISTENCIA, un **1** indicará que el alumno fue correctamente identificado como asistente a la clase y un **0** indicará que no asistió a la clase o bien, que debido a una incongruencia con los datos registrados por el alumno en la plataforma y en el nombre utilizado en la aplicación Zoom no fue posible hacer coincidir los valores, por lo que **la cifra que muestra el programa es una aproximación y se deberán revisar manualmente los alumnos con inasistencia** para una correcta evaluación.

![Exportacion](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEh3-4bffxjq33261EMUzKb13m2Ed0-n2kr2pwQpp3i9C7YI25KkZlJQ1BV-tyEbtCb8fCQbTV_YZOT-4N9OlSZduxxfizcrCk1gx479XYGWruTdYO_JUuNCyJfsjPmyBLg9He5a7WceYfKSMvHJ41qsmZJrL0X-jd-ly9wNX3Y8thloTOZxoOrda9VKW5E7/s969/Screenshot_6.png)

Si lo desea podrá utilizar el botón de **Exportar Resultados** para exportar tanto el listado de Resultados como el de la asistencia a la clase de Zoom / Plataforma, solo deberá asignar una ubicación y un nombre de archivo para proceder correctamente con la exportación.

![ExportacionTerminada](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEgMmLt9sEYr3Yg_zY0Tah7WylnKD8JKz7mU-yzzA0UMs4NKgadoLjZe371Cj3fIUyrhStRDDvpUnh_AiHi7F0pV1uiwyYCUTyVyGi7Ge45F9siQE7NlAnl3hZZWiIQtEiIJ54q9x1UHX9icUxmLY06PA6ek4IDSioznMrpCjIfcVvRtGCiRUUQLj06SvCS7/s967/Screenshot_7.png)

Finalmente podrá realizar una evaluación manual utilizando Excel para garantizar la integridad de los resultados

![ExportacionTerminada](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEgG8aOfIizsbQRLDChlR_-YVDOkgtnnvtF3caj_VZ3FexR0Yaes8uKym8PayhAfzIlVZuYccd-9QZtNW6loyGj0Fnc_AJOe45bSMNK-jnsqgShoWXzW5ClVxr0TMb-81rD2Q9FokyXXSW_r_vKlWDyjW5RW6Tb9yEOJmN12RfqMHdGhcciJ94A0W-7TBXnM/s829/Screenshot_8.png)

![ExportacionTerminada2](https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEi37pXnlEW6EUnimfN4Xiy3qV6o1C1iMw4eQWlshMOYVd6kyNwkJFzkhDSr80910x28NrG7WrXGl2ANK5RDQemLe_Oa6ZMeVQ1sIppUZIv0WYirjwE-o8Q_WNHD70WkAS_YJs0FnU_SzbKtobqCh6x_yEARsjjqoaBYWJE7wSbeh3b2CzvEm5VCDdYFUIod/s832/Screenshot_9.png)

