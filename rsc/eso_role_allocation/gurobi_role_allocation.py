# -*- coding: utf-8 -*-
"""
Created on Tue Nov 10 14:08:21 2020

@author: Maximiliano Escobar Viegas
"""
import pandas as pd
import numpy as np
import gurobipy as gb
from gurobipy import *
import csv
import os


availability = pd.read_excel (r'C:/Users/benja/Desktop/Sisua Digital SpA/PROYECTS/ESO Chile/eso_role_allocation/base.xlsx',"availability").fillna(0)
certified= pd.read_excel (r'C:/Users/benja/Desktop/Sisua Digital SpA/PROYECTS/ESO Chile/eso_role_allocation/base.xlsx',"roles").fillna(0)
matrix= pd.read_excel (r'C:/Users/benja/Desktop/Sisua Digital SpA/PROYECTS/ESO Chile/eso_role_allocation/base.xlsx',"main_matrix").fillna(0)

matrix=matrix.melt(id_vars=["worker", "role"],var_name="day",value_name="available")
matrix['day'] = matrix['day'].astype(int)
#print(" ")
#print("-------Main Matrix---------")
print(matrix)


#Set of days
days=np.arange(1,len(availability.columns),1) 
#print(" ")
#print("-------Set of days---------")
#print(days)

#Set of workers
workers= (availability.loc[:,'worker']).values
#print(" ")
#print("--------Set of workers----------")
#print(workers)

#Set of availability
availability=availability.melt(id_vars=["worker"],var_name="day",value_name="availability")
#print(" ")
#print("---------Set of availability------------")
#print (availability)

#Set of roles certification
certified=certified.melt(id_vars=["worker"],var_name="roles",value_name="certified")
#print(" ")
#print("----------Set of roles-----------")
#print(certified)

roles=certified['roles'].unique()
roles_weight=[1,1,1,1,1,1,1,1,1]
roles_total=len(roles)
R=tupledict()

for i in range(len(roles)):
    r= roles[i]
    w=roles_weight[i]
    R[(i)]=w

#Availability
A=tupledict()
for i in range(len(matrix)):
    w= int(matrix.iloc[i]["worker"])
    r= int(matrix.iloc[i]["role"])
    d= int(matrix.iloc[i]["day"])
    a= int(matrix.iloc[i]["available"])
    A[(w,r,d)]=a

#Model creation
m= gb.Model("ESO")

#Decision variables
H=m.addVars(len(workers),len(roles),len(days),vtype=GRB.BINARY,name="handover")
X=m.addVars(len(workers),len(roles),len(days),vtype=GRB.BINARY,name="allocation")
S= m.addVars(len(workers),len(roles),len(days),vtype=GRB.BINARY,name="roleshift")
MIN= m.addVars(len(workers),vtype=GRB.INTEGER,name="min_allocation")
MAX= m.addVars(len(workers),vtype=GRB.INTEGER,name="max_allocation")


#Constraints

#Total of roles constraint
constr_total_roles= m.addConstrs(gb.quicksum(X[w,r,d] for w in range(len(workers)) for r in range(len(roles))) \
                      == roles_total for d in range(len(days)))

#One role a day constraint
constr_one_role_a_day=m.addConstrs(gb.quicksum(X[w,r,d] for r in range(len(roles)))<= 1 for w in range(len(workers)) \
                                   for d in range(len(days)))

#Allocate only if available
constr_only_if_available=m.addConstrs(X[w,r,d]<= A[(w,r,d)] for w in range(len(workers)) for r in range(len(roles)) \
                                     for d in range(len(days))) 

#No more than 3 TCO's allocation per worker

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==1:
            m.addConstr(gb.quicksum(X[w,r,d] + H[w,r,d] for d in range(len(days))) <=3)


#Always 1 role of each kind allocated per day

constr_one_role_of_each_per_day= m.addConstrs(gb.quicksum(X[w,r,d] for w in range(len(workers))) <=1 \
                                              for r in range(len(roles)) for d in range(len(days)))

##MIP Constraint MIN relation
constr_min= m.addConstrs(MIN[w]<=gb.quicksum(X[w,r,d] for r in range(len(roles)) for d in range(len(days)))\
                         for w in range(len(workers)))

#MIP Constraint MAX relation
constr_max=m.addConstrs(gb.quicksum(X[w,r,d] for r in range(len(roles)) for d in range(len(days)))<=MAX[w] \
                        for w in range(len(workers)))
  
    
# Activate H only if available    
constr__activate_H_only_if_available=m.addConstrs(H[w,r,d]<= A[(w,r,d)] for w in range(len(workers)) for r in range(len(roles)) \
                                     for d in range(len(days))) 


# H availability constraint

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==0 or r==1 or r==2:
            for d in range(len(days)):
                if d >= 1:
                    m.addConstr(X[w,r,d] <= A[(w,r,d-1)])
                    

# H activation constraint

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==0 or r==1 or r==2:
            for d in range(len(days)):
                if d >= 1:
                    m.addConstr(H[w,r,d-1] == S[w,r,d])

    
# S Border condition

for w in range(len(workers)):
    for r in range(len(roles)):
        for d in range(len(days)):
            if d==0:
                m.addConstr(X[w,r,d]==S[w,r,d])
  
                
# S activation constraint

for w in range(len(workers)):
    for r in range(len(roles)):
        for d in range(len(days)):
            if d!= 0:
                m.addConstr(S[w,r,d] >= X[w,r,d] - X[w,r,d-1])            
   
                             
# Consecutive role allocation for TCO,COE,VLTI               

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==0 or r==2:
            for d in range(len(days)):
                if d+3<=days[len(days)-1]:
                    m.addConstr(S[w,r,d] <= X[w,r,d+1])
                    m.addConstr(S[w,r,d] <= X[w,r,d+2])


# Consecutive role allocation for ut1.ut2,ut3,ut4,vst,vista               

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==1 or r==3 or r==4 or r==5 or r==6 or r==7 or r==8:
            for d in range(len(days)):
                if d+2<=days[len(days)-1]:
                    m.addConstr(S[w,r,d] <= X[w,r,d+1])

                
# Maximun consecutive role allocation for COE and VLTI

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==0 or r==2:
            for d in range(len(days)-5):
                m.addConstr(X[w,r,d] + X[w,r,d+1] + X[w,r,d+2] + X[w,r,d+3] +X[w,r,d+4] + X[w,r,d+5] <=5) 


# Only 1 shift before minimum amount of alocations for COE and VLTI

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==0 or r==2:
            for d in range(len(days)-3):
                m.addConstr(S[w,r,d] + S[w,r,d+1] + S[w,r,d+2] <=1) 
                

# Only 1 shift before minimum amount of alocations for TST(UT1, UT2, UT3, UT4, VST, VISTA)

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==3 or r==4 or r==5 or r==6 or r==7 or r==8:
            for d in range(len(days)-2):
                m.addConstr(S[w,r,d] + S[w,r,d+1] <=1) 
                

# Maximun consecutive role allocation for TST(UT1, UT2, UT3, UT4, VST, VISTA)

for w in range(len(workers)):
    for r in range(len(roles)):
        if r==3 or r==4 or r==5 or r==6 or r==7 or r==8:
             for d in range(len(days)-4):
                m.addConstr(X[w,r,d] + X[w,r,d+1] + X[w,r,d+2] + X[w,r,d+3] +X[w,r,d+4] <=4)                


# Head of MSE's occupation

for w in range(len(workers)):
        for r in range(len(roles)):
            if r==0 and w==0:
                m.addConstr(gb.quicksum(X[w,r,d] for d in range(len(days)))<= gb.quicksum(X[w,r,d]\
                            for w in range(len(workers)) for d in range(len(days)))*0.45,name="MSI_occupation" )
                
                                        
# Objective function
#obj= gb.quicksum(MAX[w] for w in range(len(workers)))-gb.quicksum(MIN[w] for w in range(len(workers)))

obj= gb.quicksum(X[w,r,d]*R[r] for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days)))\
                -(gb.quicksum(MAX[w] for w in range(len(workers)))-gb.quicksum(MIN[w] for w in range(len(workers))))*100000 \
                + gb.quicksum(S[w,r,d] for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days)))\
                + gb.quicksum(X[w,r,d]  for w in range(1) for r in range(1) for d in range(len(days)))*10

m.setObjective(obj,GRB.MAXIMIZE)
m.update()


#Optimize call
m.optimize()


# Write output to csv
var_names = []
var_values = []

for var in m.getVars():
    if var.X >0: 
        var_names.append(str(var.varName))
        var_values.append(var.X)


try:
    os.remove('model_output.csv')
except:
    print("No file to delete")    

with open('model_output.csv', 'w') as myfile:
    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL)
    wr.writerows(zip(var_names, var_values))

    


#for w in range(len(workers)):
##    aux_matrix=matrix.loc[matrix["worker"]== w]
#    for r in range(len(roles)):
##        aux_matrix=aux_matrix.loc[aux_matrix["role"] == r]
#        for d in range(len(days)):
##            aux_matrix=aux_matrix.loc[aux_matrix["day"] == d]
#           if r==0 and d==0:
#               print(X[(w,r,d)])
#            


    
            
    