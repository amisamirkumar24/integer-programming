
"""-------
Created on 20/07/2021
By Devendra Singh
"""
from ortools.linear_solver import pywraplp
import pandas as pd

# 1. Loading the input data 
sup_stock = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Supplier stock", index_col=0)
raw_material_cost = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Raw material costs", index_col=0)
raw_material_shipping = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Raw material shipping", index_col=0)
production_req = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Product requirements", index_col=0)
production_capacity = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Production capacity", index_col=0)
customer_demand = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Customer demand", index_col=0)
production_cost = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Production cost", index_col=0)
shipping_costs = pd.read_excel("Data_sc_network_optimization.xlsx", sheet_name = "Shipping costs", index_col=0)

# Changing NaN values to 0 for easy computation
customer_demand = customer_demand.fillna(0)
production_req = production_req.fillna(0)
sup_stock = sup_stock.fillna(0)
production_capacity = production_capacity.fillna(0)
raw_material_cost = raw_material_cost.fillna(0)
production_cost = production_cost.fillna(0)




# Getting list of factories 

factories = list(raw_material_shipping.columns )
print("Factories:\n",factories)

# Getting list of materials
materials = list(raw_material_cost.columns)
print("Materials: \n",materials)

# Getting list of suppliers
suppliers = list(raw_material_cost.index)
print("Suppliers: \n",suppliers)

#Getting list of products
products = list(production_req.index)
print("Products: \n",products)

#Getting list of customers
customers = list(customer_demand.columns)
print("Customers: \n",customers)



# 2. Creating ortools solver to solve integer programming (SCIP used to solve mixed integer linear programming)

solver = pywraplp.Solver_CreateSolver('SCIP')
# Adding decision variable to solver with name ex-Factory A_Materiaal A_Supplier A, total 60 variable

orders = {}
for factory in factories:
    for material in materials:
        for supplier in suppliers:
            orders[(factory, material, supplier)] = solver.IntVar(0, solver.infinity(), 
                                          factory+"_"+material+"_"+supplier)

# decision vriable for production volume, total 12 variable
production_volume = {}
for factory in factories:
    for product in products: 
        production_volume[(factory, product)] = solver.IntVar(0, solver.infinity()  , factory+"_"+product)

# Decision variable for delivery quantity, total 48 variable
delivery = {}
for factory in factories: 
    for customer in customers:
        for product in products: 
            delivery[(factory, customer, product)] = solver.IntVar(0, solver.infinity(), factory+"_"+customer+"_"+product)



# C. Defining the constraints that ensure factories produce more than they ship to the customers

for product in products: 
    for factory in factories:
        c = solver.Constraint(0, solver.infinity())
        c.SetCoefficient(production_volume[(factory, product)] , 1)
        for customer in customers:             
            c.SetCoefficient(delivery[(factory, customer, product)], -1)


# D. Defining the constraints that ensure that customer demand is met(delivery is greater tha or equal to customer demand)

for customer in customers: 
    for product in products:
        
        c = solver.Constraint(int(customer_demand.loc[product][customer]),solver.infinity())
        for factory in factories: 
            c.SetCoefficient(delivery[(factory,customer,product)], 1)

# E. Defining the constraints that ensure that suppliers have all ordered items in stock     

for supplier in suppliers: 
    for material in materials: 
        c = solver.Constraint(0, int(sup_stock.loc[supplier][material]))
        for factory in factories: 
            c.SetCoefficient(orders[(factory, material, supplier)],1)
            
            
# F. Defining the constraints that ensure that factories order enough material to be able to manufacture all items

for factory in factories:
    for material in materials:
        c = solver.Constraint(0,solver.infinity())
        for supplier in suppliers:
            c.SetCoefficient(orders[(factory, material, supplier)],1)
            for product in products:
                c.SetCoefficient(production_volume[(factory, product)], - production_req.loc[product][material])
            

# G.Defining the constraints that ensure that the manufacturing capacities are not exceeded
                 
for factory in factories: 
    for product in products: 
        c = solver.Constraint(0, int(production_capacity.loc[product][factory]))  
        c.SetCoefficient(production_volume[(factory, product)],1)

# H.  Defining the objective function.

cost = solver.Objective()

# Material Costs  + shipping costs 
for factory in factories: 
    for supplier in suppliers:
        for material in materials:
            cost.SetCoefficient(orders[(factory, material, supplier)] , 
                                       raw_material_cost.loc[supplier][material] + raw_material_shipping.loc[supplier][factory])
            

# production cost of each factory 
for factory in factories: 
    for product in products: 
        cost.SetCoefficient(production_volume[(factory, product)], int(production_cost.loc[product][factory]))

# shipping cost to customers 
for factory in factories: 
    for customer in customers:
        for product in products: 
            cost.SetCoefficient(delivery[(factory, customer, product)], int(shipping_costs.loc[factory][customer]))


# I. Solving the ILP and determine the optimal overall cost
            
cost.SetMinimization()
status = solver.Solve()

if status == solver.OPTIMAL:
    print("Optimal Solution Found")
print("Optimal Overall Cost: ", solver.Objective().Value())

# printig supplier's Bill for all factories also priting  value of material delivered
print("\nSupplier Bill and order quantity")
print("****************************")
for factory in factories:
    print(factory,":")
    
    for supplier in suppliers:
        factory_cost = 0
        print("  ",supplier,":")
        for material in materials:
            print("\t",material,":", orders[(factory, material, supplier)].solution_value())
            
            factory_cost += orders[(factory, material, supplier)].solution_value() * raw_material_cost.loc[supplier][material]
            factory_cost += orders[(factory, material, supplier)].solution_value() * float(raw_material_shipping.loc[supplier][factory])
        print("  ",supplier," Bill: ", factory_cost)

#  printing manufacturing Cost for all factories
print("Production Volume:")
print("****************************")

for factory in factories:
    print(factory,":")
    production_cost_total = 0
    for product in products:
        if production_volume[(factory, product)].solution_value() >0:
            print("  ",product,": ",production_volume[(factory, product)].solution_value())
            production_cost_total += production_volume[(factory, product)].solution_value() * production_cost.loc[product][factory]
    print("   Manufacturing cost: ", production_cost_total)


# printing shipping Cost
print("\nShipping to Customer:") 
print("****************************")

for customer in customers:   
    shipping_cost = 0
    print(customer)
    for product in products:
        print("  ", product)
        for factory in factories: 
            print("\t",factory,": ",delivery[(factory, customer, product)].solution_value())
            shipping_cost += delivery[(factory, customer, product)].solution_value() * shipping_costs.loc[factory][customer]
    print("   Shipping Cost: ", shipping_cost)




#  final output of next few lies is cost incured(per unit) to satisfy demand of customer for a product
print("\nMaterial Bifurcation and Cost per unit")
print("****************************")

for customer in customers:
    print(customer)
    for product in products:
        
        unit_cost_per_product = 0
        if int(customer_demand.loc[product][customer]) >0:
            print("  ", product)
            for factory in factories:
                
                if delivery[(factory, customer, product)].solution_value() >0:
                    print("\t", factory, ": ")
                    # Calculating the Shipping cost from factory to customer based on number of products
                    shipping_cost = delivery[(factory, customer, product)].solution_value() * shipping_costs.loc[factory][customer]
                    # Calculating the manufacturing cost 
                    man_cost = delivery[(factory, customer, product)].solution_value() * production_cost.loc[product][factory]
                    unit_cost_per_product += shipping_cost
                    unit_cost_per_product += man_cost
                    material_cost_to_customer = 0
                    for material in materials:
                        material_units = 0
                        material_units += delivery[(factory, customer, product)].solution_value() * production_req.loc[product][material]
                        
                        print("\t  ",material,": ", material_units)  
                        #raw material cost
                        material_cost = 0
                        #raw material cost
                        rshipping_cost = 0 
                        material_count = 0
                        for supplier in suppliers:
                            material_cost +=  orders[(factory, material, supplier)].solution_value() * raw_material_cost.loc[supplier][material]
                            rshipping_cost += orders[(factory, material, supplier)].solution_value() * raw_material_shipping.loc[supplier][factory]
                            material_count += orders[(factory, material, supplier)].solution_value()
                        material_cost_to_customer = ((material_cost + rshipping_cost)/material_count) * material_units
                        unit_cost_per_product += material_cost_to_customer
            print("\t cost per unit : ", unit_cost_per_product/int(customer_demand.loc[product][customer]))


    

# calculating total cost to factories for satisfying customer demand
print("---------------------------------------")
print("Total cost to factories and units delivered")
for factory in factories:
    print("    "+factory)
    factory_cost = 0
    volume_produced = 0
    for supplier in suppliers:
        for material in materials:
            factory_cost = factory_cost + orders[(factory,material,supplier)].solution_value()*(raw_material_shipping.loc[supplier][factory] + raw_material_cost.loc[supplier][material])
    for product in products:
        factory_cost = factory_cost + production_volume[(factory,product)].solution_value()*production_cost.loc[product][factory]
        for customer in customers:
            factory_cost = factory_cost + delivery[(factory,customer,product)].solution_value()*shipping_costs.loc[factory][customer] 
            volume_produced = volume_produced + delivery[(factory,customer,product)].solution_value()
    print("        Cost:"+str(factory_cost)+"    Units Delivered:"+str(volume_produced))      
    




