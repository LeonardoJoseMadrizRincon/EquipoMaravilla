import math
import pandas

#Lectura de Archivo excel a trabajar
filePath = "./valores_ascensor.xlsm"
df = pandas.read_excel("valores_ascensor.xls")

# Esta función se encarga de realizar los calculos expuesta en el reglamento covenin 621 -3
def calculus(n): 
# Variables de entradas a tantear
    z= float(df.iloc[n,1])  #Número de ascensores-->Excel       
    p= float(df.iloc[n,2])  #Capacidad nominal-->Excel
    vn= float(df.iloc[n,3]) #Velocidad nominal-->Excel


# Cuentas obtenidas de los datos de cada grupo de ascensores
    na = float(df.iloc[n,9])    #Pisos totales encima de la plataforma princial-->Excel
    ne = float(df.iloc[n,4])    #Pisos no atendidos-->Excel
    ep = 3.5                    #Altura entre pisos
    ns = na - ne                #Pisos recorridos
    B = ns*35 + 3*15            #Poblacion total
    phi = 1                     #Aceleración


#Valores pre-definidos por las reglas covenin, segun tablas
    pv = float(df.iloc[n,5])    #Personas por viaje-->Excel/Tabla 5
    np = float(df.iloc[n,6])    #Número de paradas-->Excel/Tabla 3
    t1 = float(df.iloc[n,7])    #Tiempo -->Excel/Tabla 7
    t2 = float(df.iloc[n,8])    #Excel/Tabla 8


#Calculo de otras variables
    ha = na * ep #Recorrido superior total
    he = round(ne* ep,4) # recorrido expreso
    hs = round(ha - he,4) # Recorrido sobre la planta principal con servicios de ascensores
    check = round(float(math.sqrt(hs*(phi)/np)),4)


# Calculo de los tiempos de un viaje completo
    if (check >= vn) and (df.iloc[n,0] != "Grupo A"):
        tvc = 2*(ha/vn)+(vn/phi+t1)*(np+1)-hs/(np*vn)+t2*pv

    elif (check < vn) and (df.iloc[n,0] != "Grupo A"):
        tvc =2*(ha/vn)-hs/vn+2*(vn/phi)+(2*hs/(np*check))*(np-1) + t1*(np+1) + t2*pv
 
    elif (check >= vn) and (df.iloc[n,0] == "Grupo A"):
        tvc = 2*ha/vn + (vn/phi + t1)*(np + 1) + t2*pv

    elif (check <= vn) and (df.iloc[n,0] == "Grupo A"):
        tvc = 2*ha/check+vn/phi+ha/vn+t1*(np+1)+t2*pv
     
    else:
        print("Hay algo raro con la velocidad nominal")

    ta = (1/10)*tvc #Tiempo adicional
    ttv = tvc + ta #Tiempo de un circuito completo
    i = ttv/z
    c = (300*pv*(z*100))/(ttv*816)
    return (c,i)


def report(c,i,grupo,w):
    with open('./report_result.txt',w,encoding="utf8") as f:
        if (c < 12) or (i>40):
            f.write(f"{grupo}:\n")
            f.write(f"\t c = {c} %\n")
            f.write(f"\t i = {i} sg\n")
            f.write(f"\tLos resultados no son admitibles\n\n")

        elif (c>12) or (i<49):
            f.write(f"{grupo}:\n")
            f.write(f"\t c = {c} %\n")
            f.write(f"\t i = {i} sg\n")
            f.write(f"\tLos resultados cumplen con los requisitos del tráfico vertical\n\n")

        elif (c<12) or (i<49):
            f.write(f"{grupo}:\n")
            f.write(f"\t c = {c} %\n")
            f.write(f"\t i = {i} sg\n")
            f.write(f"\tEl valor de la capacidad de transporte no se encuentra en norma\n\n")

        elif (c>12) or (i>49):
            f.write("grupo:\n")
            f.write(f"\t c = {c} %\n")
            f.write(f"\t i = {i} sg\n")
            f.write(f"\tEl valor del intervalo probable no se encuentra en norma\n\n")


def run():
    c1, i1 = calculus(0)
    report(c1,i1,str(df.iloc[0,0]),"w")


    c2, i2 = calculus(1)
    report(c2,i2,str(df.iloc[1,0]),"a")


    c3, i3 = calculus(2)
    report(c3,i3,str(df.iloc[2,0]),"a")

    print("El programa ha finalizado con éxitos!")

if __name__ == "__main__":
    run()
