import os
import openpyxl as opxl
import Funciones as fs
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys

path = 'ezsave.xlsx'
wb = opxl.load_workbook(path)

print("---------------------------------")
print("-------Bienvenido a EZSave-------")
print("---------------------------------")
print()
print("Presiona Enter para continuar.")
enter = input()

#Menu principal
print("------------Menú---------------")
print()
print("Iniciar sesión................1")
print()
print("Crear un usuario..............2")
print()
print("Salir.........................3")
print()
print("Elige una opción:")
op = int(input())


while op < 1 or op > 3:
    print("Esa no es una opción válida, por favor selecciona una opción valida.")
    op = int(input())
    print()
if op == 1: #Iniciar Sesión
    w2 = fs.ISV()
    wb = w2
elif op == 2: #Crear Usuario
    w2 = fs.MenuCU()
    wb = w2
    w2 = fs.ISV()
    print()
elif op == 3:
    print("Decidiste salir.")

