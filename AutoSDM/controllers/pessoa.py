class Pessoa(object):
    
    '''def __init__(self, idd, matricula, nome):
        self.idd = idd
        self.matricula = matricula
        self.nome = nome
'''
    def setIdd(self, idd):
        self.idd= idd

    def setMatricula(self, matricula):
        self.matricula = matricula

    def setNome(self, nome):
        self.nome = nome

    def getIdd(self, idd):
        return self.idd

    def getMatricula(self, matricula):
        return self.matricula

    def getNome(self, nome):
        return self.nome

    idd = property( fget = getIdd, fset = setIdd)
    matricula = property( fget = getMatricula, fset = setMatricula)
    nome = property( fget = getNome, fset = setNome)
