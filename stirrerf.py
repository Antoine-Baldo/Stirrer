from pywinauto import application
import time
import pyautogui
import os

# Number of cycles 
# You can change the cycle number
n = 1

# Setpoint temperatures 
# You can add or change the temperature (between 25 and 350 celsuis degree)
T=n*[49,50]

# Time steps delta_t
# You can add or change the time (in second) between two temperature
delta_t = n*[15,15]

# Setpoint speed
# You can change the stirrer speed (between 100 and 1500 revolution per minutes)
speed=n*len(T)*[300]

# Setpoint deltat
# You can change the error for the delta time (in celsuis degree)
e = 1

# Setpoint port
# You can change the port number (between 5 and 8)
port = 6

# Open The Stirrer software
app = application.Application()
app.start(r'"C:\Program Files (x86)\Stirrer Software\MSUserSoft.exe"')

def change_temperature(temperature):
	pyautogui.moveTo(1200,680)
	pyautogui.doubleClick()
	time.sleep(1)
	for letter in str(temperature):
		pyautogui.press(letter)

def change_temperature(temperature, speed):
	pyautogui.moveTo(1200,680)
	pyautogui.doubleClick()
	time.sleep(.5)
	# write digits for temperature
	for letter in str(temperature):
		pyautogui.press(letter)
	# Moves to stirrer
	pyautogui.press('tab')
	# Write digits for speed
	for letter in str(speed):
		pyautogui.press(letter)

def typing(app,variable):
	print variable
	# Convert to string
	str_variable = str(variable)
	print str_variable
	# Digit each number one by one
	for letter in str_variable:
		print letter, type(letter)
		pyautogui.press(str_variable)

# Message box
pyautogui.alert(text="Make sure that the hot plate is pluged in the good port and don't touch the mouse after pressing START", title='Warning', button='START')

# Move to the port selection
pyautogui.press('tab')
pyautogui.press('tab')

save = 0

# Select the port wanted
for i in range(port-1):
	pyautogui.press('down') 

pyautogui.press('enter')

# Pause during 7 second because the software take some times to pop up and we have to wait for it
time.sleep(7)
# Dummy value to initate loop
current_temp = 0
for i in range(len(T)):	

	# Change temperature and speed
	change_temperature(T[i], speed[i])

	# Click on Start to start the temperature changed
	pyautogui.moveTo(1200,735)
	pyautogui.click()

	# Save the Data as an xls excel files named 'temp-test.xls' on the Desktop
	time.sleep(2)
	pyautogui.moveTo(1200,790)
	pyautogui.click()
	app.SaveAs.edit.SetText('temp-test.xls')
	app.SaveAs.Save.Click()
	pyautogui.moveTo(1000,535)
	pyautogui.click()
	
	# Begin the temperature testing for each temperature
	while ((float(current_temp))<(T[i]-e) or ((T[i]+e)<(float(current_temp)))):
		# Recover the data from the Excel files named temp-test.xls
		excel = open('temp-test.xls','r')
		data = excel.read()
		current_temp = data.replace('\t',' ').replace('\n',' ').split(' ')[-2]
		excel.close()
		time.sleep(3)
		# Save the Data as an xls excel files named 'temp-test.xls' on the Desktop)
		pyautogui.moveTo(1200,790)
		pyautogui.click()
		app.SaveAs.edit.SetEditText('temp-test.xls')
		app.SaveAs.Save.Click()
		pyautogui.moveTo(1000,535)
		pyautogui.click()

	print 'Temperature %d celsius degree OK!' % T[i]
	print 'Timer of %d second begin' % delta_t[i]
	save = save +1

	# Pause wanted
	for j in xrange(delta_t[i],0,-1):
		time.sleep(1)
		print j
	
	print 'Timer of %d second is done' % delta_t[i]

	#Click on stop so that we can change temperature again
	pyautogui.moveTo(1389,735)
	pyautogui.click()
	
	# Save the Data as an xls excel files named 'HotPlate.txt' on the Desktop
	pyautogui.moveTo(1200,790)
	pyautogui.click()
	app.SaveAs.edit.SetText('HotPlate %d.txt' % save)
	app.SaveAs.Save.Click()

# Save all the excel file in one big file
f_output = open('HotPlateFinal.txt', 'w')
f_output.write('Time(s)\tSpeed(1/min)\tTemp(C)\tCycle Number\n')

prev_step_time = 0
for i in range(1,save+1):
	f_input = open('HotPlate %d.txt' % i,'r')
	content = f_input.readlines()
	counter = 0
	for line in content:
		components = line.replace('\n','').split('\t')
		if counter != 0:
			time = int(components[0]) + prev_step_time
			line_input = str(time)
			for j in range(1,len(components)):
				line_input += '\t' + components[j]
			line_input += '\t' + '%d'% i + '\n' 
			f_output.write(line_input)
		else:
			counter += 1
	prev_step_time = time
	f_input.close()
f_output.close()

# Delete all the useless files
for i in range(1,save+1):
	os.remove('HotPlate %d.txt' % i)

# Close the The Stirrer software
pyautogui.moveTo(1389,790)
pyautogui.click()
