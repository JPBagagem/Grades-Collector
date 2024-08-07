class Student:
    def __init__(self, nMec, name):
        self.__nMec = nMec
        self.__name = name
        self.__notes = []
        self.secundario = "-"

    def __str__(self):
        return f"{self.__nMec} - {self.__name}, {self.__notes}; Nota Secundário: {self.secundario}; {self.getMedia()}"

    def getNMec(self):
        return self.__nMec

    def getName(self):
        return self.__name
    
    def getNote(self, subject):
        note = self.searchNote(subject)
        if note == None:
            return 0
        return note.getValue()
        
    def searchNote(self, subject):
        for note in self.__notes:
            if note.getSubject() == subject:
                return note
        return None
    
    def getRecursos(self):
        total = 0
        for note in self.__notes:
            total += note.getRecursos()
        return total
    
    def addNote(self, subject, note):
        self.__notes.append(Note(subject, note))

    def addRecurso(self, nRecursos, subject):
        self.searchNote(subject).addRecurso(nRecursos)

    def wasUpped(self, subject):
        return self.searchNote(subject).getRecursos() > 0

    def getMedia(self):
        return sum([note.getValue() for note in self.__notes]) / len(self.__notes)
    

class Note:
    def __init__(self, subject, value):
        self.__subject = subject
        self.__value = value
        self.__recursos = 0

    def __str__(self):
        return f"{self.__subject} : {self.__value}, Nº de recursos: {self.__recursos}"

    def getSubject(self):
        return self.__subject

    def getValue(self):
        return self.__value
    
    def getRecursos(self):
        return self.__recursos

    def addRecurso(self, nRecursos):
        self.__recursos += nRecursos