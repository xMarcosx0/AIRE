stateDiagram-v2
    [*] --> Asignado
    Asignado --> En_Inspeccion
    En_Inspeccion --> Inspeccion_Completada
    En_Inspeccion --> Inspeccion_Observaciones
    Inspeccion_Completada --> En_Revision
    Inspeccion_Observaciones --> En_Revision
    En_Revision --> Aprobado
    En_Revision --> Requiere_Correccion
    Requiere_Correccion --> En_Inspeccion
    Aprobado --> [*]