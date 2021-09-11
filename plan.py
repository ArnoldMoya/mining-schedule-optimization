# -*- coding: utf-8 -*-
"""
Created on Mon Sep  6 20:32:46 2021

@author: Lenovo
"""

import pyomo.environ as pyo
import pandas as pd

# CARGA DE DATOS
# Requerimiento
requerimiento = pd.read_excel('D:/COMACSA/Planificación/requerimiento.xlsx', 
                            index_col = 0, header = 0,
                            sheet_name="requerimiento")

# Capacidades teóricas
capacidades = pd.read_excel('D:/COMACSA/Planificación/capacidades.xlsx', 
                            index_col = 0, header = 0)
# Tiempo máximo de uso de molinos
Tmax = 30 #días
# Indices
N = list(capacidades.index.map(str)) # productos
M = list(capacidades.columns.map(str)) #molinos

# Cij: capacidad teorica del producto i en el molino j
C = {(r,c):capacidades.at[r,c] for r in N for c in M}

# di: demanda del producto i
d = {r:requerimiento.at[r,'Demanda'] for r in N}

# Verificamos si existen productos demandados con capacidad teórica igual a 0
for p in N:
    if max([C[p,m] for m in M]) ==0 and d[p]>0:
        print("Imposible",p,d[p])
        d[p] = 0

# MODELADO
model = pyo.ConcreteModel()

# tij = t del prod i en el molino j
model.t = pyo.Var(N,M,within=pyo.NonNegativeIntegers) #Integers NonNegativeReals

# Objetivo: minimizar el tiempo total de todos los productos en todos los 
# molinos
def obj_rule(mdl):
    return sum(mdl.t[p,m] for p in N for m in M)
model.objetivo = pyo.Objective(rule=obj_rule,sense=pyo.minimize)

# Restricción: Producción mayor a la demanda
def demanda_rule(mdl,p):
    return (d[p],sum(C[p,m]*mdl.t[p,m]*1/3 for m in M),None)
model.cumplir_demanda = pyo.Constraint(N,rule=demanda_rule)

# Restricción: Tiempo máximo de trabajo por molino
def tiempo_rule(mdl,m):
    return (0,sum(mdl.t[p,m] for p in N),Tmax)
model.tiempo_max = pyo.Constraint(M,rule=tiempo_rule)

# Resolver problema
opt = pyo.SolverFactory('glpk')
result_obj = opt.solve(model, tee=True)

# Escribir en archivo Excel
plan = [(p,m,pyo.value(v)) for (p,m),v in model.t.items()]
df_plan = pd.DataFrame(plan,columns=["producto","molino","días"])
df_plan.to_excel("D:/COMACSA/Planificación/plan optimizado v3.xlsx")



