#-*- coding: utf-8 -*-

#=====================Librerias==============================#
import os, sys,subprocess,inspect, time ,exceptions,shutil, arcpy
# ------------------------------------------------------------
try:
    #=========Funciones Auxiliares=====================#
    def getPythonPath():
        pydir = sys.exec_prefix
        pyexe = os.path.join(pydir, "python.exe")
        if os.path.exists(pyexe):
            return pyexe
        else:
            raise RuntimeError("python.exe no se encuentra instalado en {0}".format(pydir))

    def directorioyArchivo ():
        archivo=inspect.getfile(inspect.currentframe()) # script filename
        directorio=os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe()))) # script directory
        return archivo, directorio

    def copia_archivo(origen,destino):
        if os.path.exists(origen):
            with open(origen, 'rb') as forigen:
                with open(destino, 'wb') as fdestino:
                    shutil.copyfileobj(forigen, fdestino)

    #=========Validación de requerimientos=====================#

    pyexe = getPythonPath()
    if "x64" in r"%s"%(pyexe):
        pyexe=pyexe.replace("ArcGISx64","ArcGIS")

    # ------------------------------------------------------------
    archivo,directorio =directorioyArchivo()
    ruta = os.getcwd() + os.sep
    ruta_lectura=directorio+"\\lista_archivos.txt"
    home =os.path.expanduser("~")
    ruta_copia=home+"\\"+"Documents\ArcGIS\AddIns\Desktop"+str(arcpy.GetInstallInfo()['Version'])





    txt= open(ruta_lectura,"r")
    archivos =txt.readlines()

    for archivo in archivos:
        if "\n" in archivo:
            p,f=os.path.split(archivo.replace("\n",""))
        else:
            p,f=os.path.split(archivo)
        copia_archivo(directorio+"\\"+archivo.replace("\n",""),ruta_copia+"\\"+f)

    scriptAuxiliar="instalador_auxiliar.py"
    verPythonfinal= pyexe

    verPythonDir=pyexe.replace("\\python.exe","")
    script=directorioyArchivo()
    script1=script[1]+"\\"+scriptAuxiliar
    script=script[1]+"\\"+"get-pip.py"
    comando=r"start %s %s"%(pyexe,script)
    comando1=r"start %s %s"%(pyexe,script1)

    ff=subprocess.Popen(comando,stdin=None,stdout=subprocess.PIPE,shell=True,env=dict(os.environ, PYTHONHOME=verPythonDir))
    astdout, astderr = ff.communicate()
    hh=subprocess.Popen(comando1,stdin=None,stdout=subprocess.PIPE,shell=True,env=dict(os.environ, PYTHONHOME=verPythonDir))
    astdout, astderr = hh.communicate()

except exceptions.Exception as e:
    print e.__class__, e.__doc__, e.message
    os.system("pause")