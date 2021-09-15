def myfunc(skyString):
    outputString = '       '
    for i in range(1,len(skyString)):
        print(i)
        if i % 2 == 0:
            outputString[i] = skyString[i].upper()
        else:
            outputString[i] = skyString[i].lower()
    return skyString

testicles = 'MOTHER FUCKING TESTICLE FUCKER'
print(myfunc(testicles))