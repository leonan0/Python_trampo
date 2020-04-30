severidade = {
    1 : 'Planejada',
    2 : 'Baixa',
    3 : 'Media',
    4 : 'Alta',
    5 : 'Critica'
}

status = {
    1 : 'Aberto',
    2 : 'Aguardando cliente',
    3 : 'Em atendimento',
    4 : 'Aguardando execução de JOB',
    5 : 'Em homologação',
    6 : 'Aguardando RDM',
    7 : 'Cancelado',
    8 : 'Aguardando fornecedor',
    9 : 'Resolvido',
    10 : 'Fechado',
}


def getkey(dicionario, param):
    x = ''
    for (key,value) in dicionario.items():
        if value == param:
            x = key
            return key

