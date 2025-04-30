import requests

code = '''
stateDiagram-v2
    [*] --> Asignado
    Asignado --> En_Inspeccion
    En_Inspeccion --> Inspeccion_Completada
    En_Inspeccion --> Inspeccion_Observaciones
    Inspeccion_Completada --> En_Revision
    En_Revision --> Aprobado
'''

response = requests.post(
    "https://kroki.io/mermaid/png",
    data=code.encode("utf-8"),
    headers={"Content-Type": "text/plain"}
)

with open("diagrama.png", "wb") as f:
    f.write(response.content)

print("Imagen guardada como diagrama.png")
