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

num_decades = 12 # From 1m to 100M ohm

PREC_TABLE= ["1", "5"] # Percentual values

POWER_TABLE= ["1/10", "1/8", "1/4"] # Watts values

FOORPRINT_REF_TABLE= ["0402", "0603", "0805", "1206"] # Foot values

lib_ref = 'RES' # Sch lib

import pyodbc

# Connect to DB
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Resistor_SMD')

# Print existing data in the DB   
for row in cursor.fetchall():
    print (row)

#pos_prec_table =0 
#pos_power_table = 0
pos_foot_table = 0
id_DB = 1

foot1_ref = FOORPRINT_REF_TABLE[pos_foot_table]
foot2_ref = FOORPRINT_REF_TABLE[pos_foot_table+1]
foot3_ref = FOORPRINT_REF_TABLE[pos_foot_table+2]
foot4_ref = FOORPRINT_REF_TABLE[pos_foot_table+3]

#Sample insert
for pos_power_table in range(len(POWER_TABLE)):
    for pos_prec_table in range(len(PREC_TABLE)):
        for res_scale_rage in range(num_decades-1):
            for loop_res_value in range (len(E96_RES_TABLE)):
                print(id_DB)
                #cursor.execute("insert into Resistor_SMD([Id],[Value]) values(?, 1)", (data, ))
                #cursor.execute("insert into Resistor_SMD([Id],[Comment]) values(?, ?)", (data, E96_RES_TABLE[id_loop]))
                
                #res_value = (round(E96_RES_TABLE[loop_res_value]*10**res_scale_rage,3))

                # 1m, 10m, 100m (1m to 999m)
                if res_scale_rage < 3: 
                    res_value_str = str(round(E96_RES_TABLE[loop_res_value]*10**res_scale_rage,3)) + 'm'
                    
                # 1, 10, 100 (1 to 999R)
                elif res_scale_rage >= 3 and res_scale_rage < 6: 
                    res_value_str = str(round(E96_RES_TABLE[loop_res_value]*10**(res_scale_rage-3),3)) + 'R'

                # 1k, 10k, 1000k (1k to 999k)
                elif res_scale_rage >= 6 and res_scale_rage < 9:  
                    res_value_str = str(round(E96_RES_TABLE[loop_res_value]*10**(res_scale_rage-6),3)) + 'k'

                # 1M, 10M, 100M (1M to 100M)
                elif res_scale_rage >= 9 and res_scale_rage < 12:  
                    res_value_str = str(round(E96_RES_TABLE[loop_res_value]*10**(res_scale_rage-9),3)) + 'M'
                    
                else:
                    res_value_str = str(x)

                make_description = 'RES SMD ' + res_value_str + ' \u03A9 ' + '\u00B1' + PREC_TABLE[pos_prec_table] + '% ' + POWER_TABLE[pos_power_table] + 'W'
                cursor.execute("insert into Resistor_SMD([Id],[Description],[Comment],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3],[Footprint Ref 4]) values(?, ?, ?, ?, ?, ?, ?, ?)", (id_DB, make_description, res_value_str, lib_ref, foot1_ref, foot2_ref, foot3_ref, foot4_ref))
                id_DB=id_DB+1


# 0R
res_value_str = str(0)
for pos_power_table in range (len(POWER_TABLE)):
    print(id_DB)         
    make_description = 'RES SMD ' + res_value_str + 'R \u03A9 ' + POWER_TABLE[pos_power_table] + 'W'
    cursor.execute("insert into Resistor_SMD([Id],[Description],[Comment],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3],[Footprint Ref 4]) values(?, ?, ?, ?, ?, ?, ?, ?)", (id_DB, make_description, res_value_str, lib_ref, foot1_ref, foot2_ref, foot3_ref, foot4_ref))
    id_DB=id_DB+1
    
# 100M
res_value_str = str(100) + 'M'
for pos_power_table in range(len(POWER_TABLE)):
    for pos_prec_table in range(len(PREC_TABLE)):
        print(id_DB)
        make_description = 'RES SMD ' + res_value_str + ' \u03A9 ' + '\u00B1' + PREC_TABLE[pos_prec_table] + '% ' + POWER_TABLE[pos_power_table] + 'W'
        cursor.execute("insert into Resistor_SMD([Id],[Description],[Comment],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3],[Footprint Ref 4]) values(?, ?, ?, ?, ?, ?, ?, ?)", (id_DB, make_description, res_value_str, lib_ref, foot1_ref, foot2_ref, foot3_ref, foot4_ref))
        id_DB=id_DB+1

# Commint to DB
cursor.commit()

#


# Print existing data in the DB   
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Resistor_SMD')
for row in cursor.fetchall():
    print (row)













    

# https://datatofish.com/how-to-connect-python-to-ms-access-database-using-pyodbc/

