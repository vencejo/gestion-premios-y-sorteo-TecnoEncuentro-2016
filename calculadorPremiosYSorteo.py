import openpyxl
import pprint
import datetime
import random

class Equipo():
    
    def __init__(self, workBook, nombre):
        self.equipo = workBook.get_sheet_by_name(nombre)
        self.nombre = nombre
        self.premiosGanados = []
        self.integrantes = []
        self.integrantes.append(self.equipo['B5'].value) #Tutor
        for i in range(16):
            integrante = self.equipo['A' + str(10 + i)].value
            if integrante == None:
                break
            else:
                self.integrantes.append(integrante)
                
    def guardaPremio(self,premio):
        self.premiosGanados.append(premio)
    
    def imprimePremio(self, premio):
        print ("{}-> {} = {} ".format(self.nombre ,premio, getattr(self,premio)))
                
               
class EquipoCompeticion(Equipo):
    
    def __init__(self, workBook, nombre):
        super().__init__(workBook, nombre)
        self.mostrarPuntos = False
        self.puntos = self.equipo['F2'].value + self.equipo['F3'].value + self.equipo['F4'].value 
        self.innovacion = self.equipo['F9'].value
        self.cultura = self.equipo['F10'].value
        self.matematicas = self.equipo['F11'].value
        self.especial = self.equipo['F12'].value
        self.gotzilla = self.equipo['F13'].value
        self.hormiga = self.equipo['F14'].value
        self.troll = self.equipo['F15'].value
        self.taller = self.equipo['F16'].value
        self.milojos = self.equipo['F17'].value
        self.picasso = self.equipo['F18'].value
        self.plomo = self.equipo['F19'].value
        self.pluma = self.equipo['F20'].value
        self.colmena = self.equipo['F21'].value
        self.promesas = self.equipo['F22'].value
        minutos1 = self.equipo['H2'].value.minute
        segundos1 = self.equipo['H2'].value.second
        t1 = datetime.timedelta(minutes= minutos1, seconds = segundos1) 
        minutos2 = self.equipo['H3'].value.minute
        segundos2 = self.equipo['H3'].value.second
        t2 = datetime.timedelta(minutes= minutos2, seconds = segundos2) 
        minutos3 = self.equipo['H4'].value.minute
        segundos3 = self.equipo['H4'].value.second
        t3 = datetime.timedelta(minutes= minutos3, seconds = segundos3) 
        self.tiempo = t1 + t2 + t3
        self.mejorRecoleccion = self.calculaMejorPrueba( prueba = "recoleccion")
        self.mejorLogistica = self.calculaMejorPrueba( prueba = "logistica")
        
    def calculaMejorPrueba(self, prueba = "recoleccion"):
        mejorPuntuacion = 0
        mejorTiempo = datetime.timedelta(minutes= 0, seconds = 0) 
        for i in range(3):
            tipoPrueba = (self.equipo['G'+str(i+2)].value).lower()
            puntos = self.equipo['F'+str(i+2)].value
            minutos = self.equipo['H'+str(i+2)].value.minute
            segundos = self.equipo['H'+str(i+2)].value.second
            tiempo = datetime.timedelta(minutes= minutos, seconds = segundos) 
            if  prueba == "recoleccion" and tipoPrueba == 'r':
                if  (puntos == mejorPuntuacion and tiempo < mejorTiempo) or \
                     ( puntos > mejorPuntuacion):
                    mejorPuntuacion, mejorTiempo = puntos, tiempo
            if  prueba == "logistica" and tipoPrueba == 'l':
                if  (puntos == mejorPuntuacion and tiempo < mejorTiempo) or \
                     ( puntos > mejorPuntuacion):
                    mejorPuntuacion, mejorTiempo = puntos, tiempo
        
        return (mejorPuntuacion, mejorTiempo)
  
        
    def __str__(self):
        if self.mostrarPuntos:
            return "{}-> puntos = {} , tiempo = {}".format(self.nombre , self.puntos, self.tiempo) 
        else:
            return "{} ".format(self.nombre)
        
    
    def __lt__(self, otroEquipo):
        """ Ordenacion por puntos, pero si dos equipos tienen los mismos puntos, ordenarlos 
        por tiempo"""
        if self.puntos == otroEquipo.puntos:
            # El que tenga mayor tiempo quedara por detras en la clasificacion
            return self.tiempo > otroEquipo.tiempo  
        return self.puntos < otroEquipo.puntos
    
    
class EquipoExhibicion(Equipo):
    
    def __init__(self, workBook, nombre):
        super().__init__(workBook, nombre)
        self.seccion = "exhibicion"
        self.puntos = equipo['F1'].value
        self.innovacion = equipo['F4'].value
        self.cultura = equipo['F5'].value
        self.matematicas = equipo['F6'].value
        self.especial = equipo['F7'].value
        self.mostrarPuntos = False
        
    def __str__(self):
        if self.mostrarPuntos:
            return "{}-> puntos = {}".format(self.nombre , self.puntos) 
        else:
            return "{} ".format(self.nombre) 
       
    def __lt__(self, otroEquipo):
        """ Ordenacion por puntos, pero si dos equipos tienen los mismos puntos, ordenarlos 
        por tiempo"""
        return self.puntos < otroEquipo.puntos
    
    
    

# ------------------------------- Gestión de premios --------------------------------------

# Creando las listas de equipos
listaEquiposCompeticion = []
listaEquiposExhibicion = []

wb = openpyxl.load_workbook('tecnoencuentro.xlsx')
for nombre in wb.get_sheet_names():
    equipo = wb.get_sheet_by_name(nombre)
    if equipo['A1'].value == "Nombre Equipo":
        if equipo['D1'].value == 'Exhibición':
            listaEquiposExhibicion.append(EquipoExhibicion(wb, nombre))
        else:
            listaEquiposCompeticion.append(EquipoCompeticion(wb, nombre))

def imprime(texto, listaEquipos, atributo1=None, atributo2=None):
	print("-"*30)
	print(texto)
	print("")
	for index, eq in enumerate(listaEquipos):
		print(eq)
		if atributo1:
			if atributo1 == "mejorRecoleccion" :
				print(atributo1 + "-> puntos = " + str(eq.mejorRecoleccion[0]), ", mejor tiempo = " + str(eq.mejorRecoleccion[1])) 
			elif atributo1 =="mejorLogistica":
				print(atributo1 + "-> puntos = " + str(eq.mejorLogistica[0]), ", mejor tiempo = " + str(eq.mejorLogistica[1])) 
			else:
				print(atributo1 + " -> " + str(getattr(eq,atributo1)))
		if atributo2:
			print(atributo2 + " -> " + str(getattr(eq,atributo2)))
		print("")
		if index == 2: 
			break
	print("-"*30)
	

    

def imprimePodium(tipoPrueba="competicion"):
    """ Imprime los tres primeros"""
    
    print("PODIUM de " + tipoPrueba)
    print("="*50)
    
    if tipoPrueba == "competicion":
        listaEquipos = listaEquiposCompeticion       
    elif tipoPrueba == "exhibicion":
        listaEquipos = listaEquiposExhibicion
        
    # Ordeno los equipos según su puntuacion
    listaEquipos.sort(reverse=True)
    # Guardo los premios en la memoria de los 3 equipos ganadores
    for i, eq in enumerate(listaEquipos):
        eq.guardaPremio(tipoPrueba + " : " + str(i+1) + " puesto")
        if i == 2:
            break
    # Imprimo el podium
    if tipoPrueba == "competicion":
        imprime("Podium de " + tipoPrueba, listaEquipos, atributo1="puntos",atributo2="tiempo")
    elif tipoPrueba == "exhibicion":
        imprime("Podium de " + tipoPrueba, listaEquipos, atributo1="puntos")
        

    
def imprimeMejoresRecolectores():
    
    print("MEJORES RECOLECTORES de la competicion")
    print("="*50)
    
    # Ordeno los equipos según su recoleccion
    listaEquiposCompeticion.sort(key= lambda eq: (eq.mejorRecoleccion[0],
                                                  1/(eq.mejorRecoleccion[1].seconds+1)),
                                                  reverse=True)
    # Guardo el premio en la memoria del mejor recolector
    listaEquiposCompeticion[0].guardaPremio("Mejor Recolector")
    # Imprimo los mejores recolectores
    imprime("Mejores Recolectores:", listaEquiposCompeticion, atributo1="mejorRecoleccion")
    
     
def imprimeMejoresLogisticos():
    
    print("MEJORES LOGISTICOS de la Competicon")
    print("="*50)
    # Ordeno los equipos según su recoleccion
    listaEquiposCompeticion.sort(key= lambda eq: (eq.mejorLogistica[0],
                                                  1/(eq.mejorLogistica[1].seconds+1)),
                                                  reverse=True)
    # Guardo el premio en la memoria del mejor logistico
    listaEquiposCompeticion[0].guardaPremio("Mejor Logistico")
    # Imprimo los mejores logisticos
    imprime("Mejores Logisticos:", listaEquiposCompeticion, atributo1="mejorLogistica")
    
    
def imprimePremiosEspeciales(seccion = "competicion"):
    
    print("PREMIOS ESPECIALES de " + seccion)
    print("="*50)
    if seccion == "competicion":
        listaEquipos = listaEquiposCompeticion
        premiosEspeciales = ["gotzilla","hormiga","troll","taller","milojos","picasso",
                             "plomo","pluma","colmena","promesas","innovacion", "cultura", 
                             "matematicas","especial",]
        
    if seccion == "exhibicion":
        listaEquipos = listaEquiposExhibicion
        premiosEspeciales = ["innovacion", "cultura", "matematicas","especial"]
            
    for premio in premiosEspeciales:
        # Ordeno los equipos según el premio
        listaEquipos.sort(key= lambda eq: getattr(eq,premio), reverse=True)
        
        # Guardo el premio en la memoria del ganador del mismo
        type(listaEquipos[0])
        listaEquipos[0].guardaPremio(premio)

        imprime("Mejores " + premio + " " + seccion, listaEquipos, atributo1=premio)
        print("")
        
        
def imprimeCuadroDeHonor(seccion = "competicion"):
	
	if seccion == "competicion":
		listaEquipos = listaEquiposCompeticion
		
	if seccion == "exhibicion":
		listaEquipos = listaEquiposExhibicion
	
	print("")
	print("*-*-"*20)
	print( "Cuadro de honor del TecnoEncuentro".upper())
	print("Seccion: " + seccion)
	print("-"*20)
	
	listaEquipos.sort(key= lambda eq: len(eq.premiosGanados), reverse=True)
	for eq in listaEquipos:
		print("")
		print(eq.nombre)
		print("Premios Ganados: {}".format(len(eq.premiosGanados)))
		for premio in eq.premiosGanados:
			print(premio + ", ", end="")
		print("")
	print("")
    
def imprimeSorteo(numPremiosSorteo = 20):
	print("")
	print("*-*-"*20)
	print( "AGRACIADOS DEL SORTEO".upper())
	print("-"*20)
	#Genero la lista de participanes que son subceptibles de entrar en el sorteo,
	# Es decir, su equipo no debe de haber quedado en el podium de la competicion
	listaParticipantesDelSorteo = []
	listaEquipos= listaEquiposCompeticion + listaEquiposExhibicion  
	equiposConPremio = []
	for equipo in listaEquipos:
		for premio in equipo.premiosGanados:
			if premio.startswith("competicion"):
				equiposConPremio.append(equipo)           
	equiposSorteo = set(listaEquipos) - set(equiposConPremio)
	integrantesYaAnotados = []
	for equipo in equiposSorteo:
		for integrante in equipo.integrantes:
			if integrante not in integrantesYaAnotados:	
				listaParticipantesDelSorteo.append((integrante, equipo.nombre))
				integrantesYaAnotados.append(integrante)
	# Se barajan las cartas
	random.shuffle(listaParticipantesDelSorteo)
	for i in range(numPremiosSorteo):
		print(str(i+1) + "- " + str(listaParticipantesDelSorteo[i]))


if __name__ == "__main__":
	
	print("")
	print("")
	print("#"*72)
	print("				TECNOENCUENTRO ALMERIA 2016 ")
	print("#"*72)
	print("")
	print("")
	imprimePremiosEspeciales(seccion="exhibicion") 
	print("-*"*40)
	print("")
	imprimePodium(tipoPrueba="exhibicion")
	print("-*"*40)
	print("")
	imprimePremiosEspeciales(seccion="competicion") 
	print("-*"*40)
	print("")
	imprimeMejoresLogisticos()
	print("-*"*40)
	print("")
	imprimeMejoresRecolectores()
	print("-*"*40)
	print("")
	imprimePodium(tipoPrueba="competicion")
	print("-*"*40)
	print("")
	imprimeCuadroDeHonor(seccion = "competicion")
	print("-*"*40)
	print("")
	imprimeCuadroDeHonor(seccion = "exhibicion")
	imprimeSorteo()
	
