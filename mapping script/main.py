
#ensayo general

#abre archivo y lee y escribe linea por linea en la variable lines
with open('test.txt', encoding='utf8') as f:
   lines = f.readlines()



data_maestra = ["JAULA,RESERVA,SKU,PRODUCTO,CANTIDAD,ESTADO RESERVA,DESP,BLQ,COMUNA,DIRECCION,CLIENTE,TIPO DE RESERVA,PESO,VOLUMEN,PEC,PKT,CANAL,NºGUIA,ORIGEN"]


counter = 1
currentJaula_lista=[]
linea_reserva_lista=[]
disponible = []


#while desde contador 1 hasta la longitud total del archivo
while counter < len(lines):
    #define una linea current_line que toma el valor de la linea correspondiente al contador
    current_line = lines[counter]
    
    #si la linea que está evaluando empieza en Jaula hace un remplazo de parametros para dejar solo el numero de la jaula en la variable currentJaula_lista
    if current_line.startswith("Jaula"):
        jaula = current_line.replace("\n","").replace(" ","").replace("Jaula:","").split()
        currentJaula_lista = jaula 
    #si empieza por reserva hace un remplazo para cambiar , por . ademas de cambiar por "" los caracteres y palabras mencionadas    
    elif current_line.startswith("Reserva:"):
        filter_line = current_line.replace("Reserva:","").replace("Canal:","").replace(",",".").replace("\n","").replace("(","").replace(")","").lstrip()
        #hace split de la data para agregarlo a una lista dataSeparated con un split de parametro tres espacios
        dataSeparated = filter_line.split("  ")
        
        final_list =  []

        for x in dataSeparated:
            x = x.strip().strip(".").strip()
            final_list.append(x)

        
        lista_sin_espacios=[]

        for x in final_list:
            if x[0:1].isalnum():
                lista_sin_espacios.append(x)

        if len(lista_sin_espacios) == 6:
            lista_sin_espacios[3:5]=[" ".join(lista_sin_espacios[3:5])]

        if len(lista_sin_espacios) == 7:
            lista_sin_espacios[3:6]=[" ".join(lista_sin_espacios[3:6])]
    
        linea_reserva_lista = lista_sin_espacios
    
    elif current_line.startswith("Tienda"):
        filter_line = current_line.replace("Tienda Venta:","").replace("Obs:","").replace("\n","").lstrip()
        dataSeparated = filter_line.split("  ")
        final_list =  []
        for x in dataSeparated:
            if x[1:2].isalnum():
                final_list.append(x)
        if len(final_list) < 2:
            final_list.append("")   
    elif current_line.lstrip().startswith(("D","C/R","C/D","R")):  
        filter_line = current_line.replace("\n","").replace(",",".").replace("*","")     
        dataSeparated =filter_line.split("  ")   
        if len(dataSeparated)> 2:
            lista_sin_espacios=[]
            for x in dataSeparated:
                x = x.strip()
                lista_sin_espacios.append(x)  

            final_list =  []
            for x in lista_sin_espacios:
                if x[0:1].isalnum():
                    final_list.append(x)
                disponible = final_list
            
            if disponible[1] == "214" or disponible[1] == "221" or disponible[1] == "262" or disponible[1] == "508" or disponible[1] == "367":
                disponible.insert(1,"")

            if len(disponible) > 13:
                disponible[5:7]=[" ".join(disponible[5:7])]    

            if len(disponible[8]) == 6:
                if disponible[8][1] =="." and disponible[8][5] == ".":
                    disponible[8] = disponible[8].replace(".","")
            
           # lista_completa = currentJaula_lista  + linea_reserva_lista + disponible 
            lista_completa = []
            lista_completa = currentJaula_lista + linea_reserva_lista[0:1] + disponible[4:6]  + disponible[7:8] + disponible[6:7] + disponible[3:4] + disponible[12:13] +linea_reserva_lista[2:5] +disponible[0:1] +disponible[8:12] + linea_reserva_lista[1:2] +disponible[1:3]
            string_madre = ",".join(lista_completa)
            data_maestra.append(string_madre)


    counter += 1    


with open('output.csv','w') as p:
    p.write('\n'.join(data_maestra))





















#limpiar una linea especifica, pasarla a un string y luego pushearla a una lista


"""
linea_random = "      Reserva: 243642434       Canal: 33    RANCAGUA        MANUEL ANTONIO MATTA 197 197                                              Manu Valenzuela"
filter_line = linea_random.replace("Reserva:","").replace("Canal:","").lstrip()
dataSeparated = filter_line.split("  ")

final_list =  []
for x in dataSeparated:
    if x[0:1].isalnum():
        final_list.append(x)
           
final_String =",".join(final_list)    

final_list1 = []

final_list1.append(final_String)

print(final_list1)
"""







# forma para pushear n jaulas a una lista
"""
counter = 0
lista= []
lista_hija = []
while counter < len(lines):
    currentLine = lines[counter]
    
    if currentLine.startswith("Jaula"):
        
        jaula = currentLine.replace("Jaula:", "").replace(" ", "").replace("\n", "")
        currentJaula = jaula 
        lista_hija.append(currentJaula)
    
    counter += 1 
lista.append(lista_hija)            
#print(lista)       
"""