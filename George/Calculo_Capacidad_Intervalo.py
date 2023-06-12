
'''  
Altura entre pisos= 3.5m
El número de piso viene dado por:
-sótanos: 3 pisos para todos
-Nro de Pisos superiores:

Edicaciones pisos>=30, hoteles
Edificaciones pisos<30, edificios residenciales 

George Galindez: 85

-Pisos de área de ocupación de 750m2

-Personas que ocupan pisos sótanos: 15/piso
-Personas que ocupan cada nivel superior 35/piso
'''

''' 
Este edificio cuenta con 85 plantas superiores y 3 sotanos, un total de 88 plantas.

La altura entre pisos es de 3.5m  

'''

import numpy
import openpyxl

wb = openpyxl.Workbook()

def main():

    # Valores establecidos para el Sistema de Ascensores.

    Altura_entre_pisos = 3.5 #m
    Sotanos = 3
    Planta_Principal = 1
    Pisos_Superiores = 84
    Pisos_Totales = Pisos_Superiores + Planta_Principal + Sotanos

    # Cantidad de personas    

    Personas_Sotanos = 15
    Personas_P_Superiores = 35
    
    P = 22 #Capacidad nominal de la cabina

    Tiempo_Apertura_Cierre = 3.95 #s

    ######### Datos Grupo A: 
    '''
        El Grupo A atiende Pb, hasta el piso 28 y 3 Sótanos.
    '''

    Poblacion_estimada_A = 1060
    
    Nro_Ascensores_A = 6
    Velocidad_Nominal_A = 10 #m/s
    Tiempo_Entrada_Salida_A = 2 #s
    eap = 3.5

    ######### Datos Grupo B:

    '''
        El grupo B atiende Pb, Piso 29 al 57 y 3 Sótanos.
    '''
    Poblacion_estimada_B = 1060

    Nro_Ascensores_B = 6
    Velocidad_Nominal_B = 10 #m/s
    Tiempo_Entrada_Salida_B = 2

     ######### Datos Grupo C:

    '''
        El grupo C atiende Pb, Piso 57 al 84 y 3 Sótanos.
    '''
    Poblacion_estimada_C = 1060

    Nro_Ascensores_C = 6
    Velocidad_Nominal_C = 10 #m/s
    Tiempo_Entrada_Salida_C = 2

    Nro_Ascensores = Nro_Ascensores_A + Nro_Ascensores_B + Nro_Ascensores_C 

    Area_Pisos = 750 #m^2

    
    Aceleracion = 1 #m/s^2

    #----------------------------------------------------------

    '''
        El sistema de ascensores se divide en 3 grupos (A, B y C) de ascensores, con 6 ascensores cada uno
        (un total de 18 ascensores)
    '''
 
    #------------------------Cálculos del Grupo A (Planta Ppaal. piso 1 hasta 28):


    print("\n Cálculos del grupo A: \n")
    print(f"[Grupo A]La Vel. Nominal establecida es: {Velocidad_Nominal_A} [m/s]")

    # P: capacidad nominal de la cabina (personas).

    Pv_A = int((3.2/P)+(0.7*P)+0.5)

    ne_A = 56 #Numero de pisos NO servidos por encima de la planta principal

    ns_A = 28 #Numero de pisos servidos encima de la planta principal

    Np_A = ns_A*(1-(((ns_A-1)/(ns_A))**(Pv_A))) # Nro de paradas probables en los pisos superiores

    na_A = ns_A + ne_A #Número total de pisos encima de la planta principal.

    Ha_A = na_A*eap #Recorrido entre la planta principal y superior

    He_A = ne_A*eap #Recorrido expreso

    Hs_A = Ha_A - He_A #Recorrido sobre la planta principal con servicio de ascensores entre la primera y la ultima parada superior

    RVn_A = numpy.sqrt(Hs_A*Aceleracion/Np_A)

    print("[Grupo A] Referencial Vel. nominal es: " , RVn_A)

    if RVn_A >= Velocidad_Nominal_A:

        Tiempo_Viaje_Completo_A = (2*(Ha_A/Velocidad_Nominal_A))+(((Velocidad_Nominal_A/Aceleracion)+Tiempo_Apertura_Cierre)*(Np_A+1))-(Hs_A/(Np_A*Velocidad_Nominal_A))+(Tiempo_Entrada_Salida_A*Pv_A)
    
    elif RVn_A <= Velocidad_Nominal_A:

        Tiempo_Viaje_Completo_A = (2*(Ha_A/Velocidad_Nominal_A))-(Hs_A/Velocidad_Nominal_A)+(2*Velocidad_Nominal_A/Aceleracion)+(2*Hs_A/(Hs_A*Aceleracion/Np_A)**Np_A)*(Np_A-1)+Tiempo_Apertura_Cierre*(Np_A+1)+Tiempo_Entrada_Salida_A*(Pv_A)

    print("[Grupo A] Tiempo de Viaje completo: ", Tiempo_Viaje_Completo_A)

    Tiempo_Total_Viaje_A = Tiempo_Viaje_Completo_A + Tiempo_Viaje_Completo_A*(30/100)

    print("[Grupo A] Tiempo Total de Viaje: ", Tiempo_Total_Viaje_A)

    print("[Grupo A] Personas por viaje: ", Pv_A)

    C_A = (300*(Pv_A)*(Nro_Ascensores_A)*100)/(Tiempo_Total_Viaje_A*Poblacion_estimada_A)

    print("[Grupo A] Capacidad de transporte: ", C_A, "%")

    I_A = Tiempo_Total_Viaje_A/Nro_Ascensores_A

    print("[Grupo A] Intervalo probable: ", I_A , "[s] \n")


    #----------Cálculos del Grupo B (Piso 29 al 56):

    print("\n Cálculos del grupo B: \n")
    print(f"[Grupo B] La Vel. Nominal establecida es: {Velocidad_Nominal_B} [m/s]")

    Pv_B = int((3.2/P)+(0.7*P)+0.5)

    ne_B = 56 #Numero de pisos NO servidos por encima de la planta principal

    ns_B = 28 #Numero de pisos servidos encima de la planta principal

    Np_B = ns_B*(1-(((ns_B-1)/(ns_B))**Pv_B)) # Nro de paradas probables en los pisos superiores

    na_B = ns_B + ne_B #Número total de pisos encima de la planta principal.

    Ha_B = na_B*eap #Recorrido entre la planta principal y superior

    He_B = ne_B*eap #Recorrido entre la planta principal y la primera planta superior servida

    Hs_B = Ha_B - He_B #Recorrido sobre la planta principal con servicio de ascensores entre la primera y la ultima parada superior

    RVn_B = numpy.sqrt(Hs_B*Aceleracion/Np_B)

    print("[Grupo B] Referencial Vel. nominal es: " , RVn_B)

    if RVn_B >= Velocidad_Nominal_B:

        Tiempo_Viaje_Completo_B = (2*(Ha_B/Velocidad_Nominal_B))+(((Velocidad_Nominal_B/Aceleracion)+Tiempo_Apertura_Cierre)*(Np_B+1))-(Hs_B/(Np_B*Velocidad_Nominal_B))+(Tiempo_Entrada_Salida_B*Pv_B)
    
    elif RVn_B <= Velocidad_Nominal_B:

        Tiempo_Viaje_Completo_B = (2*(Ha_B/Velocidad_Nominal_B))-(Hs_B/Velocidad_Nominal_B)+(2*Velocidad_Nominal_B/Aceleracion)+(2*Hs_B/(Hs_B*Aceleracion/Np_B)**Np_B)*(Np_B-1)+Tiempo_Apertura_Cierre*(Np_B+1)+Tiempo_Entrada_Salida_B*(Pv_B)

    print("[Grupo B] Tiempo de Viaje completo: ", Tiempo_Viaje_Completo_B)

    Tiempo_Total_Viaje_B = Tiempo_Viaje_Completo_B + Tiempo_Viaje_Completo_B*(30/100)

    print("[Grupo B] Tiempo Total de Viaje: ", Tiempo_Total_Viaje_B)

    print("[Grupo B] Personas por viaje: ", Pv_B)

    C_B = (300*(Pv_B)*(Nro_Ascensores_B)*100)/(Tiempo_Total_Viaje_B*Poblacion_estimada_B)

    print("[Grupo B] Capacidad de transporte: ", C_B, "%")

    I_B = Tiempo_Total_Viaje_B/Nro_Ascensores_B

    print("[Grupo B] Intervalo probable: ", I_B , "[s]")

    #-------------------Cálculos del Grupo C (Piso 57 al 85, Pb y Sótanos):

    print("\n Cálculos del grupo C: \n")

    print(f"[Grupo C] La Vel. Nominal establecida es: {Velocidad_Nominal_C} [m/s]")

    Pv_C = int((3.2/P)+(0.7*P)+0.5)

    ne_C = 56 #Numero de pisos NO servidos por encima de la planta principal

    ns_C = 28 #Numero de pisos servidos encima de la planta principal

    Np_C = ns_C*(1-(((ns_C-1)/(ns_C))**Pv_C)) # Nro de paradas probables en los pisos superiores

    na_C = ns_C + ne_C #Número total de pisos encima de la planta principal.

    Ha_C = na_C*eap #Recorrido entre la planta principal y superior

    He_C = ne_C*eap #Recorrido entre la planta principal y la primera planta superior servida

    Hs_C = Ha_C - He_C #Recorrido sobre la planta principal con servicio de ascensores entre la primera y la ultima parada superior

    RVn_C = numpy.sqrt(Hs_C*Aceleracion/Np_C)

    print("[Grupo C] Referencial Vel. nominal es: " , RVn_C)

    if RVn_C >= Velocidad_Nominal_C:

        Tiempo_Viaje_Completo_C = (2*(Ha_C/Velocidad_Nominal_C))+(((Velocidad_Nominal_C/Aceleracion)+Tiempo_Apertura_Cierre)*(Np_C+1))-(Hs_C/(Np_C*Velocidad_Nominal_C))+(Tiempo_Entrada_Salida_C*Pv_C)
    
    elif RVn_C <= Velocidad_Nominal_C:

        Tiempo_Viaje_Completo_C = (2*(Ha_C/Velocidad_Nominal_C))-(Hs_C/Velocidad_Nominal_C)+(2*Velocidad_Nominal_C/Aceleracion)+(2*Hs_C/(Hs_C*Aceleracion/Np_C)**Np_C)*(Np_C-1)+Tiempo_Apertura_Cierre*(Np_C+1)+Tiempo_Entrada_Salida_C*(Pv_C)

    # -(Hs_A/(Np_A*Velocidad_Nominal_A))
  
    print("[Grupo C] Tiempo de Viaje completo: ", Tiempo_Viaje_Completo_C)

    Tiempo_Total_Viaje_C = Tiempo_Viaje_Completo_C + Tiempo_Viaje_Completo_C*(30/100)

    print("[Grupo C] Tiempo Total de Viaje: ", Tiempo_Total_Viaje_C)

    print("[Grupo C] Personas por viaje: ", Pv_C)

    C_C = (300*(Pv_C)*(Nro_Ascensores_C)*100)/(Tiempo_Total_Viaje_C*Poblacion_estimada_C)

    print("[Grupo C] Capacidad de transporte: ", C_C, "%")

    I_C = Tiempo_Total_Viaje_C/Nro_Ascensores_C

    print("[Grupo C] Intervalo probable: ", I_C , "[s]")

    print("\n")


    def Guardar_Calculo():

        #HojaCreada = wb.create_sheet("Cálculo", 0)

        Hoja = wb.active

        Hoja["A1"] = "Hola"

        wb.save("Prueba.xlsx")

        pass

    #Guardar_Calculo()

    
if __name__ == "__main__":
    print("Running...")
    main()
    