import re

pattern = '^CA\d{2} \d{2} \d{2,3}'
test_string = 'CA22 12 128'
result = re.match(pattern, test_string)

if result:
  print("Search successful.")
else:
  print("Search unsuccessful.")	
