
exc_Dict ={}
exc_Dict["Krall,Audrey"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
exc_Dict["Dam,Phuong My"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
exc_Dict["Lee,My Dung"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
exc_Dict["Trieu,Minh Hue"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
exc_Dict["Vaccari,Sergio"] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
exc_Dict["Lagano,Lauren"] = [1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, ""]

#print(exc_Dict)
exc_Dict["Vaccari,Sergio"][10] = exc_Dict["Vaccari,Sergio"][10] + 5
exc_Dict["Lagano,Lauren"][1] = exc_Dict["Lagano,Lauren"][1] + 5
exc_Dict["Lagano,Lauren"][12] = 115
print(exc_Dict["Vaccari,Sergio"][10])       # 5
print(exc_Dict["Lagano,Lauren"][1])         # 7
print(exc_Dict["Lagano,Lauren"][12])