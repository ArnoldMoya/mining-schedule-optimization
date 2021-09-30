# -*- coding: utf-8 -*-
"""
Created on Mon Sep 27 21:44:17 2021

@author: Lenovo
"""

import pyomo.environ as pyo
import pyomo.gdp as pyogdp
import pandas as pd
import math
from pyomo.opt.results import SolverStatus

# CARGA DE DATOS
# Requerimiento
requerimiento = pd.read_excel('D:/Planificación/requerimiento.xlsx',
                              index_col=0, header=0,
                              sheet_name="Grupo 2")

# Capacidades teóricas
capacidades = pd.read_excel('D:/Planificación/capacidades.xlsx',
                            index_col=0, header=0,
                            sheet_name="Grupo 2")


# # Disponibilidad de molinos
disponibilidad = pd.read_excel('D:/Planificación/disponibilidad.xlsx',
                               index_col=0, header=0)


# Tiempo máximo de uso de molinos
Tmax = 26    #días
# Indices
N = list(requerimiento.index.map(str)) # productos
Molinos = list(capacidades.columns.map(str)) #molinos

# Cij: capacidad teorica del producto i en el molino j
C = {(p,m):capacidades.at[p,m] for p in N for m in Molinos}

# di: demanda del producto i
d = {p:requerimiento.at[p,'Demanda'] for p in N}

# Dm: disponibilidad de molino m
D = {m: disponibilidad.at[m, 'Disponibilidad'] for m in Molinos}

# Verificamos si existen productos demandados con capacidad teórica igual a 0
for p in N:
    if max([C[p,m] for m in Molinos]) == 0 and d[p] > 0:
        print("Imposible: ",p,d[p])
        d[p] = 0

# # Habilitación de multiproductos
# P = {(p1,p2,m):0 for p1 in N for p2 in N for m in Molinos}

# rangoSobreposicion = [(p1,p2,m) 
#                       for p1 in N 
#                       for p2 in N 
#                       for m in Molinos
#                       if C[p1,m] != 0 and C[p2,m] != 0 
#                       and p1 != p2 and P[p1,p2,m]== 0
#                       and d[p1] > 0 and d[p2] > 0]

print ("INICIO DE MODELO")
# MODELADO
model = pyo.ConcreteModel()

# tij = t del prod i en el molino j
model.t = pyo.Var(N,Molinos,within=pyo.NonNegativeReals,bounds=(0,Tmax)) 
# xpm: asignación de p en m
model.x = pyo.Var(N,Molinos,within=pyo.Binary)
# STpm = inicio de p en m
#model.ST  = pyo.Var(N,Molinos,within=pyo.NonNegativeReals,bounds=(0,Tmax))
model.M = pyo.Param(initialize=10*Tmax);

#model.disjunctions = pyo.Set(initialize = rangoSobreposicion, dimen = 3)
# Objetivo: minimizar el tiempo total de todos los productos en todos los 
# molinos
def obj_rule(mdl):
    return sum(mdl.x[p,m] for p in N for m in Molinos)
model.objetivo = pyo.Objective(rule=obj_rule,sense=pyo.minimize)

# Restricción: Producción mayor a la demanda
def demanda_rule(mdl,p):
    return (d[p],sum(C[p,m]*mdl.t[p,m] for m in Molinos),None)
model.cumplir_demanda = pyo.Constraint(N,rule=demanda_rule)

# # Restricción: Límite del tiempo de inicio ST de p en m
# def ST_limit(mdl,p,m):
#     return (0,mdl.ST[p,m],D[m])
# model.limite_ST = pyo.Constraint(N,Molinos,rule=ST_limit)

# Restricción: Tiempo máximo de trabajo de p en m
def tiempo_pm_max_rule1(mdl,p,m):
    return (0, mdl.t[p,m], D[m])
model.tiempo_pm_max1 = pyo.Constraint(N,Molinos,rule=tiempo_pm_max_rule1)

# Restricción: Si xpm = 0 entonces tpm = 0
def tiempo_pm_max_rule2(mdl,p,m):
    return (None, mdl.t[p,m] - D[m]*mdl.x[p,m], 0)
model.tiempo_pm_max2 = pyo.Constraint(N,Molinos,rule=tiempo_pm_max_rule2)


# Restricción: si xpm = 1 entonces tpm>=1
def tiempo_pm_max_rule3(mdl,p,m):
    return (0, mdl.t[p,m] - mdl.x[p,m], None)
model.tiempo_pm_max3 = pyo.Constraint(N,Molinos,rule=tiempo_pm_max_rule3)

# Suma de tiempos < Tmax del molino
def tiempo_rule(mdl, m):
    return (0, sum(mdl.t[p, m] for p in N), D[m])  # D[m] Tmax
model.tiempo_max = pyo.Constraint(Molinos, rule=tiempo_rule)

# # Restricción: Tiempo final máximo de p en m
# def tiempo_fin_pm_max_rule(mdl,p,m):
#     return (0, mdl.ST[p,m] + mdl.t[p,m], D[m])
# model.tiempo_fin_pm_max = pyo.Constraint(N,Molinos,rule=tiempo_fin_pm_max_rule)


# def sobreposicion(mdl,p1, p2, m):
#     return [
#         [mdl.ST[p1,m] + mdl.t[p1,m]  <= mdl.ST[p2,m] + \
#          (2 - mdl.x[p1,m] - mdl.x[p2,m]) * mdl.M],
#         [mdl.ST[p2,m] + mdl.t[p2,m]  <= mdl.ST[p1,m] + \
#          (2 - mdl.x[p1,m] - mdl.x[p2,m]) * mdl.M]
#             ]
# model.sobreposicion_rule = pyogdp.Disjunction(rangoSobreposicion,
#                                               rule=sobreposicion)
        


# Resolver problema
pyo.TransformationFactory('gdp.bigm').apply_to(model)

#opt = pyo.SolverFactory('glpk')
#opt.options['tmlim'] = 100
opt = pyo.SolverFactory('cbc', executable='D:/Planificación/Cbc-master-win64-msvc14-md/bin/cbc.exe') #glpk ,executable='D:/COMACSA/Cbc-1.1.0-win32-msvc7/bin/cbc.exe'
#opt.options['tmlim'] = 100
opt.options['seconds'] = 100
result_obj = opt.solve(model, tee=True)

plan = [(p,
         m,
         pyo.value(model.x[p,m]),#math.ceil(pyo.value(model.x[p,m])),
         pyo.value(model.t[p,m]))#math.ceil(pyo.value(model.t[p,m])))
         for p in N
         for m in Molinos]
df_plan = pd.DataFrame(plan,columns=["producto","molino","asignado","días"])
df_plan.to_excel("D:/Planificación/test grupo 2 V02.xlsx")

#
print(pyo.value(model.objetivo))
