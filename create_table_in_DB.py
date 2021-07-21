################################################################################################
# Create resistor DB
#
# Description: Auto. create rows for SMT resistor values in the MS Access Altium DB
#
# Changelog:
# - 21/07/2021 - GBPINTARELLI deploy V1.0
################################################################################################

# Sample resistor table
E96_RES_TABLE= [1.00,	1.02,	1.05,
1.07,	1.10,	1.13,
1.15,	1.18,	1.21,
1.24,	1.27,	1.30,
1.33,	1.37,	1.40,
1.43,	1.47,	1.50,
1.54,	1.58,	1.62,
1.65,	1.69,	1.74,
1.78,	1.82,	1.87,
1.91,	1.96,	2.00,
2.05,	2.10,	2.16,
2.21,	2.26,	2.32,
2.37,	2.43,	2.49,
2.55,	2.61,	2.67,
2.74,	2.80,	2.87,
2.94,	3.01,	3.09,
3.16,	3.24,	3.32,
3.40,	3.48,	3.57,
3.65,	3.74,	3.83,
3.92,	4.02,	4.12,
4.22,	4.32,	4.42,
4.53,	4.64,	4.75,
4.87,	4.99,	5.11,
5.23,	5.36,	5.49,
5.62,	5.76,	5.90,
6.04,	6.19,	6.34,
6.49,	6.65,	6.81,
6.98,	7.15,	7.32,
7.50,	7.68,	7.87,
8.06,	8.25,	8.45,
8.66,	8.87,	9.09,
9.31,	9.53,	9.76]

num_decades = 12 # From 1m to 100M ohm

PREC_TABLE= ["1", "5"] # Percentual values

POWER_TABLE= ["1/10", "1/8", "1/4"] # Watts values

FOORPRINT_REF_TABLE= ["0603", "0805", "1206"]

lib_ref = 'RES'

# ID, Value, Description, Library Ref, Part Number, Footprint Ref, Footprint Ref 1,  Footprint Ref 2, 

import pyodbc


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Resistor_SMD')


# Print existing data in the DB   
for row in cursor.fetchall():
    print (row)

pos_prec_table =0 
pos_power_table = 0
pos_foot_table = 0
id_DB = 1


#Sample insert
for res_scale_rage in range(num_decades-1):
    for loop_res_value in range (len(E96_RES_TABLE)):
        print(id_DB)
        #cursor.execute("insert into Resistor_SMD([ID],[Value]) values(?, 1)", (data, ))
        #cursor.execute("insert into Resistor_SMD([ID],[Comment]) values(?, ?)", (data, E96_RES_TABLE[id_loop]))
        
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
            res_value_str = str(0)
        
        #res_value_str = str(res_value)

            

        
        
        make_description = 'RES SMD ' + res_value_str + ' \u03A9 ' + '\u00B1' + PREC_TABLE[pos_prec_table] + '% ' + POWER_TABLE[pos_power_table] + 'W'
        foot1_ref = FOORPRINT_REF_TABLE[pos_foot_table]
        foot2_ref = FOORPRINT_REF_TABLE[pos_foot_table+1]
        foot3_ref = FOORPRINT_REF_TABLE[pos_foot_table+2]
        cursor.execute("insert into Resistor_SMD([ID],[Description],[Comment],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3]) values(?, ?, ?, ?, ?, ?, ?)", (id_DB, make_description, res_value_str, lib_ref, foot1_ref, foot2_ref, foot3_ref))
        id_DB=id_DB+1
cursor.commit()

#


# Print existing data in the DB   
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Resistor_SMD')
for row in cursor.fetchall():
    print (row)













    

# https://datatofish.com/how-to-connect-python-to-ms-access-database-using-pyodbc/

