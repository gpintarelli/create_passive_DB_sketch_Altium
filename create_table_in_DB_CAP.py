################################################################################################
# Create capacitor DB
#
# Description: Auto. create rows for SMT resistor values in the MS Access Altium DB
# Database: ID, Value, Description, Library Ref, Part Number, Footprint Ref, Footprint Ref 1,  Footprint Ref 2,  Footprint Ref 3
# Changelog:
# - 22/07/2021 - GBPINTARELLI deploy V1.0
################################################################################################

# Sample resistor table
CAP_TABLE= [1.0,   1.2,   1.5,
            1.8,   2.2,   2.7,
            3.3,   3.9,   4.7,
            5.6,   6.8,   8.2]

num_decades = 6 # From 1p to 100n F

PREC_TABLE= ["5", "10"] # Percentual values for  1p to 100n F

PREC_TABLE_u= ["10", "20"] # Percentual values for  1p to 100n F

VOLTAGE_TABLE= ["25", "50"] # Voltage values

FOOTPRINT_REF_TABLE= ["C0402", "C0603", "C0805", "C1206"] # Foot values

TYPE_DI = ["X7R"] # Theres no loop regarding type, this is reserved for future implementation

lib_ref = 'CAP' # Sch lib

import pyodbc

# Connect to DB
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Ceramic_Capacitor_SMD')

# Print existing data in the DB   
for row in cursor.fetchall():
    print (row)

pn = 'X'
manu = 'X'

#pos_prec_table =0 
#pos_power_table = 0
pos_foot_table = 0
id_DB = 1

foot1_ref = FOOTPRINT_REF_TABLE[pos_foot_table]
foot2_ref = FOOTPRINT_REF_TABLE[pos_foot_table+1]
foot3_ref = FOOTPRINT_REF_TABLE[pos_foot_table+2]
foot4_ref = FOOTPRINT_REF_TABLE[pos_foot_table+3]

flag_type_uF = 0

#Sample insert
# pF and nF range
for pos_volt_table in range(len(VOLTAGE_TABLE)):
    for pos_prec_table in range(len(PREC_TABLE)):
        for cap_scale_rage in range(num_decades):
            for loop_cap_value in range (len(CAP_TABLE)):
                print(id_DB)
                #cursor.execute("insert into Resistor_SMD([Id],[Value]) values(?, 1)", (data, ))
                #cursor.execute("insert into Resistor_SMD([Id],[Comment]) values(?, ?)", (data, E96_RES_TABLE[id_loop]))
                
                #res_value = (round(E96_RES_TABLE[loop_res_value]*10**res_scale_rage,3))

                # 1p, 10p, 100p (1p to 999p)
                if cap_scale_rage < 3:
                    cap_value_str = str(round(CAP_TABLE[loop_cap_value]*10**cap_scale_rage,3)) + 'p'
                    
                # 1n, 10n, 100n (1 to 999n)
                elif cap_scale_rage >= 3 and cap_scale_rage < 6:
                    cap_value_str = str(round(CAP_TABLE[loop_cap_value]*10**(cap_scale_rage-3),3)) + 'n'

                ## 1u, 10u, 1000u (1k to 999u)
                #elif cap_scale_rage >= 6 and res_scale_rage < 9:
                #    res_value_str = str(round(CAP_TABLE[loop_cap_value]*10**(cap_scale_rage-6),3)) + 'u'
                    
                else:
                    cap_value_str = str(x)

                make_description = 'CAP SMD ' + cap_value_str + 'F ' + '\u00B1' + PREC_TABLE[pos_prec_table] + '% ' + VOLTAGE_TABLE[pos_volt_table] + 'V ' + TYPE_DI[0]
                cursor.execute("insert into Ceramic_Capacitor_SMD([Id],[Description],[Comment],[Part Number],[Manufacturer],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3],[Footprint Ref 4]) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (id_DB, make_description, cap_value_str, pn, manu, lib_ref, foot1_ref, foot2_ref, foot3_ref, foot4_ref))
                id_DB=id_DB+1

# uF range
for pos_volt_table in range(len(VOLTAGE_TABLE)):
    for pos_prec_table in range(len(PREC_TABLE_u)):
        for cap_scale_rage in range(2): # fixed range 1 to 99uF
            for loop_cap_value in range (len(CAP_TABLE)):
                print(id_DB)
                #cursor.execute("insert into Resistor_SMD([Id],[Value]) values(?, 1)", (data, ))
                #cursor.execute("insert into Resistor_SMD([Id],[Comment]) values(?, ?)", (data, E96_RES_TABLE[id_loop]))
                
                #res_value = (round(E96_RES_TABLE[loop_res_value]*10**res_scale_rage,3))

                # 1u, 10u, 1000u (1k to 999u)
                cap_value_str = str(round(CAP_TABLE[loop_cap_value]*10**cap_scale_rage,3)) + 'u'

                make_description = 'CAP SMD ' + cap_value_str + 'F ' + '\u00B1' + PREC_TABLE_u[pos_prec_table] + '% ' + VOLTAGE_TABLE[pos_volt_table] + 'V ' + TYPE_DI[0]
                cursor.execute("insert into Ceramic_Capacitor_SMD([Id],[Description],[Comment],[Part Number],[Manufacturer],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3],[Footprint Ref 4]) values(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (id_DB, make_description, cap_value_str, pn, manu, lib_ref, foot1_ref, foot2_ref, foot3_ref, foot4_ref))
                id_DB=id_DB+1

    

# Commint to DB
cursor.commit()

#


# Print existing data in the DB   
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Ceramic_Capacitor_SMD')
for row in cursor.fetchall():
    print (row)













    

# https://datatofish.com/how-to-connect-python-to-ms-access-database-using-pyodbc/

