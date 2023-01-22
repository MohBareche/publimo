def initiales_gest(nom):
	cap = nom.split(' ')
	init = cap[0][0] + cap[1][0]
	return init

def initiales(nom):
	cap = nom.split(' ')
	init = cap[0][0] + cap[1][0]
	return init.lower()

print(initiales('Mohamed Bareche'))
print(initiales_gest('Jérôme Vaillancourt'))