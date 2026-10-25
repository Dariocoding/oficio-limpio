# Oficio Limpio

Mini-servicio **PHP + HTML** que:

1. Recibe un PDF escaneado  
2. Pide OCR a **Stirling** (`STIRLING_URL`)  
3. Genera un **Word editable** con plantilla militar (sin firmas, sellos ni CamScanner)

## Uso local

```bash
# en esta carpeta
docker build -t oficio-limpio .
docker run --rm -p 8088:80 -e STIRLING_URL=https://pdf.klowshop.com oficio-limpio
# abrir http://localhost:8088
```

## Coolify

- Proyecto: **servicios internos**
- Build pack: **Dockerfile**
- Puerto: `80`
- Variable: `STIRLING_URL=https://pdf.klowshop.com`
- Dominio sugerido: `oficio.klowshop.com` (o el que prefieras)

## Notas

- Optimizado para oficios CEFOA / solicitud de munición (3 páginas).
- Si el OCR falla, igual genera la plantilla con valores por defecto del último formato validado.
