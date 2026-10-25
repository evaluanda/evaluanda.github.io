# Prototipo: fotos del alumnado solo en el dispositivo

Prueba si un almacén con dirección fija (`evaluanda.github.io`), cargado en un
iframe dentro de una web de Apps Script, conserva las fotos entre aperturas y
con el paso de los días, sobre todo en iPad/iPhone (Safari). Nada llega a Google.

## Cómo probarlo

1. Crea un proyecto nuevo en script.google.com (no toques la app).
2. Pega `Code.gs` y crea un archivo HTML llamado `Index` con `Index.html`.
3. Implementar → Nueva implementación → Aplicación web
   (Ejecutar como: yo · Acceso: solo yo, o tu dominio).
4. Abre la URL en el iPad, guarda 1–3 fotos y pulsa «Estado del almacén».
5. Pulsa en el iframe «Pedir almacenamiento persistente».
6. Repite en Safari normal y añadiéndola a la pantalla de inicio.

## Qué mirar

| Prueba | Bien si… |
|---|---|
| Cerrar y volver a abrir | Siguen las fotos y la misma fecha «Creado» |
| A los 8–10 días sin abrirla | Ídem (la regla de 7 días de Safari no la borró) |
| «Persistente» | `true` tras pedirlo |
| Errores | «Sin respuesta del almacén» = el navegador bloquea el iframe |
