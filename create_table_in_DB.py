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
1.54,	1,58,	1.62,
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
6.49,	6.65,	6,81,
6.98,	7.15,	7.32,
7.50,	7.68,	7.87,
8.06,	8.25,	8.45,
8.66,	8.87,	9.09,
9.31,	9.53,	9.76]

PREC_TABLE= ["1", "5"] # Percentual values

POWER_TABLE= ["1/10", "1/8", "1/4"] # Watts values

FOORPRINT_REF_TABLE= ["0603", "0805", "1206"]

lib_ref = 'RES'

# ID, Value, Description, Library Ref, Part Number, Footprint Ref, Footprint Ref 1,  Footprint Ref 2, 

import pyodbc


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Resistor_SMD_0603')


# Print existing data in the DB   
for row in cursor.fetchall():
    print (row)

pos_prec_table = 1
pos_power_table = 1
pos_foot_table = 1

#Sample insert
for id_loop_res_value in range (len(E96_RES_TABLE)):
    print(id_loop_res_value+1)
    #cursor.execute("insert into Resistor_SMD_0603([ID],[Value]) values(?, 1)", (data, ))
    #cursor.execute("insert into Resistor_SMD_0603([ID],[Comment]) values(?, ?)", (data, E96_RES_TABLE[id_loop]))
    res_value = E96_RES_TABLE[id_loop_res_value]
    make_description = 'RES SMD ' + str(res_value) + ' \u03A9 ' + '\u00B1' + PREC_TABLE[pos_prec_table] + '% ' + POWER_TABLE[pos_power_table] + 'W'
    foot1_ref = FOORPRINT_REF_TABLE[pos_foot_table]
    foot2_ref = FOORPRINT_REF_TABLE[pos_foot_table+1]
    foot3_ref = FOORPRINT_REF_TABLE[pos_foot_table+2]
    cursor.execute("insert into Resistor_SMD_0603([ID],[Description],[Comment],[Library Ref],[Footprint Ref 1],[Footprint Ref 2],[Footprint Ref 3]) values(?, ?, ?, ?, ?, ?, ?)", (id_loop_res_value+1, make_description, res_value, lib_ref, foot1_ref, foot2_ref, foot3_ref))

cursor.commit()

#


# Print existing data in the DB   
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\AlitumLibrary.accdb;')
cursor = conn.cursor()
cursor.execute('select * from Resistor_SMD_0603')
for row in cursor.fetchall():
    print (row)













    

# https://datatofish.com/how-to-connect-python-to-ms-access-database-using-pyodbc/

