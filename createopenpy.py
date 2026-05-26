import openpyxl as op

work = op.Workbook()
sheet = work.active
sheet["a1"] = "ID"
sheet["b1"] = "Product Name"
sheet["c1"] = "Quantity"
sheet["d1"] = "Price"
sheet["e1"] = "Expiry Date"
sheet["f1"] = "Type of packaging"
sheet["g1"] = "Total Price"

sheet["a2"] = "1"
sheet["a3"] = "2"
sheet["a4"] = "3"
sheet["a5"] = "4"
sheet["a6"] = "5"

sheet["b2"] = "Patty"
sheet["b3"] = "Milk"
sheet["b4"] = "Mayonais"
sheet["b5"] = "Bans"
sheet["b6"] = "oil"

sheet["c2"] = 10
sheet["c3"] = 20
sheet["c4"] = 15
sheet["c5"] = 13
sheet["c6"] = 5

sheet["d2"] = 10
sheet["d3"] = 200
sheet["d4"] = 150
sheet["d5"] = 133
sheet["d6"] = 50

sheet["e2"] = 2026
sheet["e3"] = 2027
sheet["e4"] = 2027
sheet["e5"] = 2028
sheet["e6"] = 2026

sheet["f2"] = "Box"
sheet["f3"] = "Box"
sheet["f4"] = "Galon"
sheet["f5"] = "Sachet"
sheet["f6"] = "Pack"

sheet["g2"] = 100
sheet["g3"] = 4000
sheet["g4"] = 2250
sheet["g5"] = 1729
sheet["g6"] = 250
work.save("Inventory.xlsx")