import math
import pandas
import numpy as np

#Archivo excel a trabajar
filePath = "./valores_ascensor.xlsm"
df = pandas.read_excel("valores_ascensor.xls")
print(df)

def calculus(n):
# Variables de entradas a tantear
    z= float(df.iloc[n,1])  #Excel
    print("z: ",z)
    p= float(df.iloc[n,2])  #Excel
    print("p: ",p)
    vn= float(df.iloc[n,3]) #Excel
    print("vn: ",vn)
    
# Cuentas obtenidas de los datos de cada grupo de ascensores
    na = float(df.iloc[n,9])
    ne = float(df.iloc[n,4])    #Excel
    ep = 3.5
    ns = na - ne
    print("ns: ",ns)
    print("ne: ",ne)
    B = ns*35 + 15*3
    phi = 1

#Valores pre-definidos por las reglas covenin, segun tablas
    pv = float(df.iloc[n,5])    #Excel
    print("pv: ",pv)
    np = float(df.iloc[n,6])    #Excel
    print("np: ",np)
    t1 = float(df.iloc[n,7])    #Excel
    print("t1: ",t1)
    t2 = float(df.iloc[n,8])
    print("t2: ",t2)



#Calculo de otras variables
    ha = na * ep #Recorrido superior total
    print("ha: ",ha)
    he = round(ne* ep,4) # recorrido expreso
    print("he: ",he)
    hs = round(ha - he,4) # Recorrido sobre la planta principal con servicios de ascensores
    print("hs: ",hs)
    check = round(float(math.sqrt(hs*(phi)/np)),4)
    print(check)

# Calculo de los tiempos de un viaje completo
    if (check >= vn) and (df.iloc[n,0] != "Grupo A"):
        tvc = 2*(ha/vn)+(vn/phi+t1)*(np+1)-hs/(np*vn)+t2*pv

    elif (check < vn) and (df.iloc[n,0] != "Grupo A"):
        tvc =2*(ha/vn)-hs/vn+2*(vn/phi)+(2*hs/(np*check))*(np-1) + t1*(np+1) + t2*pv
        print("TVC: ",tvc)

    elif (check >= vn) and (df.iloc[n,0] == "Grupo A"):
        tvc = 2*ha/vn + (vn/phi + t1)*(np + 1) + t2*pv
        print("TVC: ",tvc)

    elif (check <= vn) and (df.iloc[n,0] == "Grupo A"):
        print("This")
        tvc = 2*ha/check+vn/phi+ha/vn+t1*(np+1)+t2*pv
    

    else:
        print("Hay algo raro con la velocidad nominal")

    ta = (1/10)*tvc #Tiempo adicional
    print("ta: ",ta)
    ttv = tvc + ta #Tiempo de un circuito completo
    print("ttv: ",ttv)
    i = ttv/z
    c = (300*pv*(z*100))/(ttv*816)

#verificacion
    return (c,i)


def run():
    c1, i1 = calculus(0)
    print(c1)
    print(i1)
    #c2, i2 = calculus(1)
    #c3, i3 = calculus(2)
    #c4, i4 = calculus(3)

if __name__ == "__main__":
    run()
