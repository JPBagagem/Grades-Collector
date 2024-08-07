import osFuncts
import newtxt

menu = """
Escolha uma opção:

1 - Definir lista de alunos
2 - Refazer a tabela
3 - Remover um aluno
0 - Sair

Opção: """


def searchById(list, id):
    for i in range(len(list)):
        if list[i][0] == id:
            return i
    

def subjectInput(subjects):
    option = ""
    while option != 0:
        print("Escolha uma cadeira:\n")
        for i in range(len(subjects)):
            print(f"{i+1} - {subjects[i]}")
        print("0 - Voltar atrás")
        option = input("Cadeira: ")
        if not option.isdigit():
            continue
        option = int(option)
        if not (option > 0 and option <= len(subjects)):
            continue
        return subjects[option-1]

    
def yesno(prompt):
    while True:
        res = input(prompt).upper()
        if res == "S" or res == "SIM" or res == "Y" or res == "YES":
            return True
        if res == "N" or res == "NÃO" or res == "NAO" or res == "NO":
            return False


def option1(subjects, subjectspath):
    if osFuncts.file_exist_here("alunos.txt") and not yesno("Já existe uma lista de alunos deseja substituí-la?"):
        return
    subject = subjectInput(subjects)
    newtxt.defineAlunos(subject, subjectspath[subject])


def option3():
    while True:
        nMec = input("Insira o número mecanográfico do aluno que deseja retirar da lista:")
        if (len(nMec) == 5 or len(nMec) == 6) and nMec.isnumeric():
            nMec = [int(nMec)]
            break
    newtxt.removeStudent(nMec)


def main():
    option = ""
    subjects, subjectspath = osFuncts.getSubjects()
    if not osFuncts.file_exist_here("alunos.txt"):
        print("É preciso uma cadeira onde será retirada a lista de alunos:")
        option1(subjects, subjectspath)
    while option != "0":
        option = input(menu)
        match option:
            case "1":
                option1(subjects, subjectspath)
            case "2":
                newtxt.newTable(subjects, subjectspath)
            case "3":
                option3()
    

if __name__ == "__main__":
    main()