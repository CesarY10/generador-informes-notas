import pandas as pd
import sys
import shutil
import os
from docxtpl import DocxTemplate
import copy

# --- Archivo Excel con notas y datos de alumnos. ---
#para ubicar este archivo en la carpeta inputs
#r le dice a python “No interpretes los backslashes \ como caracteres especiales”.
NOTAS_ALUMNOS_PATH = r'projectExcel/inputs/Notas_Alumnos.xlsx' # -> Esto se hace para guardar el directorio sin que haya problemas


# --- lantilla Word donde se insertarán los datos. ---
PLANTILLA_CURSOS_PATH = r'projectExcel/inputs/Plantilla_Notas.docx'
#PATH para nuestra carpeta OUTPUTS
PATH_OUTPUT = r'.\outputs'
#Variable curso
CURSO = '2021/2022'

#Colores
SUSPENSO_COLOR = 'ec7c7b'
APROBADO_COLOR = 'fbe883'
NOTABLE_COLOR = '4db4d7'
SOBRESALIENTE_COLOR = '48bf91'




# -- diccionario para cambiar las asignaturas con sus tildes
dict_asig = {
    'LENGUA CASTELLANA Y LITERATURA':   'Lengua Castellana y Literatura',
    'BIOLOGIA':                         'Biología',
    'GEOGRAFIA E HISTORIA':             'Geografía e Historia',
    'MATEMATICAS':                      'Matemáticas',
    'INGLES':                           'Inglés',
    'EDUCACION FISICA':                 'Educación Física',
    'ETICA':                            'Ética',
    'CULTURA CLASICA':                  'Cultura clásica',
    'MUSICA':                           'Música',
    'TECNOLOGIA':                       'Tecnología',
    'EDUCACION PLASTICA':               'Educación Plástica',
    'FRANCES':                          'Francés',
}


# ---- DETECCION DE ERRORES ----
#obtenemos nombre de alumno y asignatura
def deteccionErrores(df):
    err1, err2, err3 = False, False, False  # creamos variables booleanas para los errores



    alumnos_list = sorted(list(df['NOMBRE'].drop_duplicates())) #Obtiene nombres de alumnos por orden alfabetico
    asignatura_list = sorted(list(df['ASIGNATURA'].drop_duplicates())) #Obtiene la asignatura por orden alfabetico

#hacemos bucle for anidado para iterar en las listas
    for al in alumnos_list:
        for asig in asignatura_list:
            filt_al_as_df = df[(df['NOMBRE'] == al) & (df['ASIGNATURA'] == asig)]
            print('')

#--- DETECCION DE ERRORES ----

#muestra error si el alumno no tiene alguna asignatura
            if(len(filt_al_as_df) == 0):
                print(f'Error: El alumno {al} no tiene la asigntura{asig} asignada')
                err1 = True  #si la funcion se cumple, entonces err1 es True

#muestra error por asignaturas repetidas
            elif(len(filt_al_as_df) > 1):
                print(f'Error: El alumno {al} tiene la asignatura {asig} repetida {len(filt_al_as_df)}')
                err2 = True  #si la funcion se cumple, entonces err2 es True

#---- error por si la nota esta fuera de rango ----
        for index, row in df.iterrows():
            trimestre_list = ['NOTA T1' , 'NOTA T2', 'NOTA T3'] #Para tener las listas de las columnas de las notas
            for trim in trimestre_list:
                if not((row[trim] >= 0.0) and (row[trim] <= 10.0)): #si la nota NO ES mayor o igual a 0.0 y menor a 10.0
                    print(f'Error: El alumno {al} tiene el campo {trim} de la asignatura {asig} fuera de rango {str(row[trim])}')
                    err3 = True  #si la funcion se cumple, entonces err3 es True

#detecta si hay algun error
        if(err1 == True) or (err2 == True) or (err3 == True):
            print('')
            print('Debes corregir los errores para continuar con la ejecucion del programa')

#cerrar sistema (importar sys)
            sys.exit(1) #1 es bueno, 0 error

#si no hay error
        else:
            print('Ningun error detectado')          



# ---- FUNCION PARA ELIMINAR LAS TILDES DE LOS TITULOS ----
def eliminarTildes(texto):
    #diccionario para cambiar las mayusculas con tildes por mayusculas sin tildes
    tildes_dict = {
        'Á' : 'A',
        'É' : 'E',
        'Í' : 'I',
        'Ó' : 'O',
        'Ú' : 'U',
    }

#lo retornamos
    textosSinTilde = texto
#funcion para retornar textos con las tildes
    for key in tildes_dict:
        textosSinTilde = textosSinTilde.replace(key, tildes_dict[key])

    return textosSinTilde


# ---- FUNCION PARA CREAR O ELIMINAR LAS CARPETAS ----
def eliminarCrearCarpetas(path):
    if os.path.exists(path): # -> Validar si existe el path
        shutil.rmtree(path) #elimina el path existente

    os.mkdir(path) #crea la carpeta 

# ---- FUNCION PARA OBTENER LAS NOTAS FINALES ----
def ObtenerNotaFinal(dict_asignatura):
    newAsignaturaDic = copy.deepcopy(dict_asignatura)
    TRIMESTRE_LIST = ['t1','t2','t3']

    #obtener nota final
    nota_media = 0
    for trim in TRIMESTRE_LIST:
        nota_media += newAsignaturaDic[trim]
    nota_media /=3
    newAsignaturaDic['nota_Final'] = round(nota_media, 1)

    #obtener calificacion
    if(nota_media < 5.0):
        calif = 'SUSPENSO'
        color_calif = SUSPENSO_COLOR
    elif (nota_media < 7.0):
        calif = 'APROBADO'
        color_calif = APROBADO_COLOR
    elif(nota_media < 9.0):
        calif = 'NOTABLE'
        color_calif = NOTABLE_COLOR
    else:
        calif = 'SOBRESALIENTE'
        color_calif = SOBRESALIENTE_COLOR

    newAsignaturaDic['calificacion'] =  calif
    newAsignaturaDic['color'] = color_calif

    return newAsignaturaDic



# ---- FUNCION PARA CREAR WORDS Y ASIGNAR TAGS ----
def crearWordAsignarTag(datos_alumnos, excel_df):
    asig_list = sorted(list(excel_df['ASIGNATURA'].drop_duplicates())) #-> lo asignamos como tipo list
    

#---- funcion para aniadir tildes a las asignaturas y no halla problema ----
    filter_td_asig = []
    for item in asig_list: 
        valorTd = dict_asig[item]
        filter_td_asig.append(valorTd.upper())
    print('')


    #---- para ordenar como lista a la columna NOMBRE ----
    nombre_Alumno_list = sorted(list(datos_alumnos['NOMBRE'])) #-> lo asignamos como tipo list

    #---- Para automatizar la iteracion de los alumnos ----
    for nombre_alumno in nombre_Alumno_list:
    #Cargar documento word
        docs_tpl = DocxTemplate(PLANTILLA_CURSOS_PATH)


        filt_datos_alumnos_df = datos_alumnos[(datos_alumnos['NOMBRE'] == nombre_alumno)]
    
    # .iloc: Se usa para recorrer celdas
        clase = filt_datos_alumnos_df.iloc[0]['CLASE']


    #Antes del contexto, pasaremos algunos valores (filtrados) al word con un bucle for
    #crear tabla de notas
        asignatura_list = [] # -> Se inicializa como una lista vacia
    #iterar sobre los indices de asignaturas
        for asig_idx in range(len(asig_list)): #la longitud de la iteracion dependera de los indices de nuestro asig_list
            asign = asig_list[asig_idx]
            filt_al_as_excel_df = excel_df[(excel_df['NOMBRE'] == nombre_alumno) & (excel_df['ASIGNATURA'] == asign)]
            print('')

    # --- diccionario de asignaturas ---
            asignatura_dict = {
                'nombre_asignatura': filter_td_asig[asig_idx],    #Clave: Asignatura, Valor: filter_td_asig(asignaturas en mayusculas) y su parametro es el indice de asig_list
                't1': round(filt_al_as_excel_df.iloc[0]['NOTA T1'],1),      #t1 tendra como valor desde el indice 0 de la columna NOTA T1
                't2': round(filt_al_as_excel_df.iloc[0]['NOTA T2'],1),      #round: sirve para redondear los decimales al momento de dar las notas
                't3': round(filt_al_as_excel_df.iloc[0]['NOTA T3'],1), 
            }
            asignatura_dict = ObtenerNotaFinal(asignatura_dict)

            asignatura_list.append(asignatura_dict) #como la lista esta avcia le agregamos el diccionario

    #se lo pasamos al contexto

    #Context (Variables en nuestro archivo word)
        context = {
            'curso': CURSO, #es un diccionario cuya llave es curso y su valor 2021/2022
            'nombre_alumno':nombre_alumno, #clave alumno, valor su nombre
            'clase': clase ,
            'asignatura_list': asignatura_list
        }

    #para renderizar nuestro contexto en el documento
        docs_tpl.render(context)
        titulo = 'NOTAS_' + nombre_alumno
        titulo = titulo.upper() #convierte el titulo en mayusculas
        titulo = eliminarTildes(titulo)
        titulo = titulo.replace(" ","_") #reemplaza los espacios en blanco con guiones bajos _
        titulo += '.docx'
        

    #para guardar nuestro documento
        docs_tpl.save(PATH_OUTPUT + '\\' + titulo)



def main():
    eliminarCrearCarpetas(PATH_OUTPUT) #llamamos a la funcion para crear/eliminar carpetas


    #Para leer Notas y Datos_Alumnos
    excel_df = pd.read_excel(NOTAS_ALUMNOS_PATH, sheet_name='Notas') #-> Para leer las notas de los alumnos en la hoja notas
                                                                     #-> el pd.read_excel recibe dos parametros, el path del archivo. y el nombre de la hoja que leer

    datos_alumnos = pd.read_excel(NOTAS_ALUMNOS_PATH, sheet_name='Datos_Alumnos') #para leer los datos en la hoja Datos_Alumnos


#llamamos a las funciones para que funcione
    deteccionErrores(excel_df)          #Deteccion de errores
    crearWordAsignarTag(datos_alumnos, excel_df)  #Creamos y asignamos tags en el word


if __name__ == '__main__':
    main()