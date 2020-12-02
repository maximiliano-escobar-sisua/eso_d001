# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 16:41:47 2020

@author: benja
"""

def mip_model():
    import pandas as pd
    import numpy as np
    import csv
    import os 
    from mip import Model, xsum, MAXIMIZE, BINARY,INTEGER, OptimizationStatus


    availability = pd.read_excel (r'C:\Users\bcerda\Documents\UiPath\p001_asignación_de_roles\rsc/eso_role_allocation/base.xlsx',"availability").fillna(0)
    certified= pd.read_excel (r'C:\Users\bcerda\Documents\UiPath\p001_asignación_de_roles\rsc/eso_role_allocation/base.xlsx',"roles").fillna(0)
    matrix= pd.read_excel (r'C:\Users\bcerda\Documents\UiPath\p001_asignación_de_roles\rsc/eso_role_allocation/base.xlsx',"main_matrix").fillna(0)## LLAMAR A optimization_base.xlsx

    matrix=matrix.melt(id_vars=["worker", "role"],var_name="day",value_name="available") # ESTO NO ES NECESARIO, USAR OPTIMIZATION_BASE EN SU FORMA LONG
    matrix['day'] = matrix['day'].astype(int)
    #print(" ")
    #print("-------Main Matrix---------")
    print(matrix)


    #Set of days
    days=np.arange(1,len(availability.columns),1) #MODIFICARLO CON LA LóGICA DE ABAJO
    #print(" ")
    #print("-------Set of days---------")
    #print(days)

    #Set of workers
    workers= (availability.loc[:,'worker']).values ## OBTIENE LOS VALORES UNICOS DE LA COLUMNA WORKERS
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
    R=dict()

    for i in range(len(roles)):
        r= roles[i]
        w=roles_weight[i]
        R[(i)]=w

    #Availability
    A=dict()
    for i in range(len(matrix)):
        w= int(matrix.iloc[i]["worker"])
        r= int(matrix.iloc[i]["role"])
        d= int(matrix.iloc[i]["day"])
        a= int(matrix.iloc[i]["available"])
        A[(w,r,d)]=a

    #Model creation
    m = Model(name = 'ESO',sense = MAXIMIZE)

    #Decision variables
    H = {(w,r,d): m.add_var(name="handover({},{},{})".format(w,r,d), var_type= BINARY)\
         for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days))}
    X = {(w,r,d): m.add_var(name="allocation({},{},{})".format(w,r,d), var_type= BINARY)\
         for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days))}
    S =  {(w,r,d): m.add_var(name="roleshift({},{},{})".format(w,r,d), var_type= BINARY)\
         for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days))}
    MIN={(w): m.add_var(name="min_allocation({})".format(w), var_type= INTEGER) for w in range(len(workers))}
    MAX={(w): m.add_var(name="max_allocation({})".format(w), var_type= INTEGER) for w in range(len(workers))}


    # Objective function
    #obj= gb.quicksum(MAX[w] for w in range(len(workers)))-gb.quicksum(MIN[w] for w in range(len(workers)))

    m.objective= xsum(X[w,r,d]*R[r] for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days)))\
                    -(xsum(MAX[w] for w in range(len(workers)))-xsum(MIN[w] for w in range(len(workers))))*100000 \
                    + xsum(X[w,r,d] for w in range(1) for r in range(1) for d in range(len(days)))*10
                    #+ xsum(S[w,r,d] for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days)))\
                        
    #Constraints

    #Total of roles constraint
    for d in range(len(days)):
        m.add_constr(xsum(X[w,r,d] for w in range(len(workers)) for r in range(len(roles))) == roles_total)

    #One role a day constraint
    for w in range(len(workers)):
        for d in range(len(days)):
            m.add_constr(xsum(X[w,r,d] for r in range(len(roles))) <= (1))
                                       

    #Allocate only if available

    for w in range(len(workers)):
        for r in range(len(roles)):
            for d in range (len(days)):
                m.add_constr(X[w,r,d] <= A[(w,r,d)])
        
    #No more than 3 TCO's allocation per worker

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==1:
                m.add_constr(xsum(X[w,r,d] + H[w,r,d] for d in range(len(days))) <=3)


    #Always 1 role of each kind allocated per day
    for r in range(len(roles)):
        for d in range(len(days)):
            m.add_constr(xsum(X[w,r,d] for w in range(len(workers))) <=1)

    ##MIP Constraint MIN relation
    for w in range(len(workers)):
        m.add_constr(MIN[w]<=xsum(X[w,r,d] for r in range(len(roles)) for d in range(len(days))))

    #MIP Constraint MAX relation
    for w in range(len(workers)):
        m.add_constr(xsum(X[w,r,d] for r in range(len(roles)) for d in range(len(days)))<=MAX[w])
      
        
    # Activate H only if available 
    for w in range(len(workers)):
        for r in range(len(roles)):
            for d in range(len(days)):
                m.add_constr(H[w,r,d]<= A[(w,r,d)])

    # H availability constraint

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==0 or r==1 or r==2:
                for d in range(len(days)):
                    if d >= 1:
                        m.add_constr(X[w,r,d] <= A[(w,r,d-1)])
                        

    # H activation constraint

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==0 or r==1 or r==2:
                for d in range(len(days)):
                    if d >= 1:
                        m.add_constr(H[w,r,d-1] == S[w,r,d])

        
    # S Border condition

    for w in range(len(workers)):
        for r in range(len(roles)):
            for d in range(len(days)):
                if d==0:
                    m.add_constr(X[w,r,d]==S[w,r,d])
      
                    
    # S activation constraint

    for w in range(len(workers)):
        for r in range(len(roles)):
            for d in range(len(days)):
                if d!= 0:
                    m.add_constr(S[w,r,d] >= X[w,r,d] - X[w,r,d-1])            
       
                                 
    # Consecutive role allocation for TCO,COE,VLTI               

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==0 or r==2 or r==1:
                for d in range(len(days)):
                    if d <=days[len(days)-3]:
                        m.add_constr(S[w,r,d] <= X[w,r,d+1])
                  

    # Consecutive role allocation for ut1.ut2,ut3,ut4,vst,vista               

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==3 or r==4 or r==5 or r==6 or r==7 or r==8:
                for d in range(len(days)):
                    if d <=days[len(days)-4]:
                        m.add_constr(S[w,r,d] <= X[w,r,d+1])
                        m.add_constr(S[w,r,d] <= X[w,r,d+2])

                    
    # Maximun consecutive role allocation for COE and VLTI after the hand over

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==0 or r==2:
                for d in range(len(days)-6):
                    m.add_constr(X[w,r,d] + X[w,r,d+1] + X[w,r,d+2] + X[w,r,d+3] +X[w,r,d+4] + X[w,r,d+5] + X[w,r,d+6] <=6)
    # en configuracion que deve ser menor e igual a 5 empieza a fallar porque no se la puede

    # Only 1 shift before minimum amount of alocations for COE and VLTI

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==0 or r==2:
                for d in range(len(days)-2):
                    m.add_constr(S[w,r,d] + S[w,r,d+1] + S[w,r,d+2] <=1 )
    #ojo cambio                 

    # Only 1 shift before minimum amount of alocations for TST(UT1, UT2, UT3, UT4, VST, VISTA)

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==3 or r==4 or r==5 or r==6 or r==7 or r==8:
                for d in range(len(days)-1):
                    m.add_constr( S[w,r,d] + S[w,r,d+1] <=1)
                    

    # Maximun consecutive role allocation for TST(UT1, UT2, UT3, UT4, VST, VISTA)

    for w in range(len(workers)):
        for r in range(len(roles)):
            if r==3 or r==4 or r==5 or r==6 or r==7 or r==8:
                 for d in range(len(days)-4):
                    m.add_constr( X[w,r,d] + X[w,r,d+1] + X[w,r,d+2] + X[w,r,d+3] +X[w,r,d+4] <=4)       


    # Head of MSE's occupation

    for w in range(len(workers)):
            for r in range(len(roles)):
                if r==0 and w==0:
                    m.add_constr(xsum(X[w,r,d] for d in range(len(days))) <= xsum(X[w,r,d]\
                                for w in range(len(workers)) for r in range(len(roles)) for d in range(len(days)))*0.45)
                    
                        
    # Objective function
    #obj= gb.quicksum(MAX[w] for w in range(len(workers)))-gb.quicksum(MIN[w] for w in range(len(workers)))

    #Optimize call
    m.emphasis = 1
    m.max_gap = 0.05
    status = m.optimize(max_seconds=300)
    if status == OptimizationStatus.OPTIMAL:
        print('optimal solution cost {} found'.format(m.objective_value))
    elif status == OptimizationStatus.FEASIBLE:
        print('sol.cost {} found, best possible: {}'.format(m.objective_value, m.objective_bound))
    elif status == OptimizationStatus.NO_SOLUTION_FOUND:
        print('no feasible solution found, lower bound is: {}'.format(m.objective_bound))
    if status == OptimizationStatus.OPTIMAL or status == OptimizationStatus.FEASIBLE:
        print('solution:')
        for v in m.vars:
           if abs(v.x) > 1e-6: # only printing non-zeros
              print('{} : {}'.format(v.name, v.x))


    # Write output to csv
    var_names = []
    var_values = []

    for v in m.vars:
        if v.x >0: 
            var_names.append(str(v.name))
            var_values.append(v.x)


    try:
        os.remove(r'C:\Users\bcerda\Documents\UiPath\p001_asignación_de_roles\rsc/eso_role_allocation/model_output.csv')
    except:
        print("No file to delete")    

    with open(r'C:\Users\bcerda\Documents\UiPath\p001_asignación_de_roles\rsc/eso_role_allocation/model_output.csv', 'w') as myfile:
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

mip_model()
        
