# BANROL SEMANAL — automatización

Mandas 1 correo con los 2 PDF → n8n te responde con los 2 Excel listos.

## Qué hace el motor
Recibe **inv705** e **inv315** (en cualquier orden, los detecta solo) y genera:
- **WEEKS OF COVERAGE**: B(MIN) C(MAX) F(ON ORDER) E(STOCK=ON HAND) desde inv705; H(WEEKLY AVERAGE = avg/4) desde inv315. D/G/I quedan como fórmulas; J/K/L fijos; A1 = fecha.
- **MIN-MAX**: reemplaza la fecha + ON HAND en la última columna (mismas columnas, no agrega).

La fecha se saca sola del "RUN DATE" del inv705.

## Archivos
- `banrol_motor.py` — el motor (CLI + microservicio)
- `templates/` — tus 2 Excel actuales, usados como plantilla (el motor solo sobrescribe las celdas que cambian)
- `requirements.txt`, `Dockerfile` — para levantar el microservicio
- `n8n_workflow.json` — flujo de n8n para importar

## Probar local (sin n8n)
```bash
pip install -r requirements.txt
python3 banrol_motor.py inv705.pdf inv315.pdf --out ./salida
```

## Levantar el microservicio (lo que llama n8n)
```bash
docker build -t banrol .
docker run -d --name banrol -p 8000:8000 banrol
# prueba:  curl -F pdf1=@inv705.pdf -F pdf2=@inv315.pdf http://localhost:8000/generate -o banrol.zip
```
Si tu n8n corre en Docker, ponlos en la misma red y usa `http://banrol:8000/generate`.

## Montar en n8n
1. Importa `n8n_workflow.json`.
2. **Gmail Trigger**: pon tu credencial Gmail. El filtro ya busca asunto `banrol semanal` con PDF sin leer.
3. **Motor (HTTP Request)**: cambia `http://TU_HOST:8000/generate` por la URL real del microservicio.
4. **Responder correo**: pon tu credencial Gmail.
5. Activa el workflow.

> Nota: nombres de los campos binarios de los adjuntos (`attachment_0`, `attachment_1`) pueden variar según tu versión de n8n — verifica el output del Gmail Trigger y ajústalos en el nodo HTTP.

## Uso semanal
1. Produces el inv705 y conviertes el inv315 (.rpt → PDF).
2. Reenvías **ambos PDF en un solo correo** con asunto `banrol semanal <fecha>`.
3. Minutos después recibes la respuesta con los 2 Excel.

## Nota sobre inv315
El promedio mensual cambia lento (usa 6 meses). Si un correo no trae inv315, esas telas mantienen su H anterior. Con mandar inv315 completo (todas las telas) se actualizan todas; el que llega hoy es solo del grupo 2560.
