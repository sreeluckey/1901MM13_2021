def regex_renamer():

	# Taking input from the user

	print("1. Breaking Bad")
	print("2. Game of Thrones")
	print("3. Lucifer")

	webseries_num = int(input("Enter the number of the web series that you wish to rename. 1/2/3: "))
	season_padding = int(input("Enter the Season Number Padding: "))
	episode_padding = int(input("Enter the Episode Number Padding: "))
     
	import re
	import os
	import shutil
	dictn={1:"Breaking Bad",2:"Game of Thrones",3:"Lucifer"}
	try:
		shutil.rmtree(f"corrected_srt\\{dictn[webseries_num]}")
	except:
		pass
	shutil.copytree(f"wrong_srt\\{dictn[webseries_num]}",f"corrected_srt\\{dictn[webseries_num]}")
	files_list=os.listdir(f"corrected_srt\\{dictn[webseries_num]}")	
	
	for file_name in files_list:
		nums_list = re.findall(r'\d+', file_name)
		words_list=re.split(r'\.', file_name)
		ep_name=words_list[0].split("-")
		if(webseries_num == 1):
			os.rename(f"corrected_srt\\{dictn[webseries_num]}\\{file_name}", f"corrected_srt\\{dictn[webseries_num]}\\{dictn[webseries_num]} Season {nums_list[0].zfill(season_padding)} Episode {nums_list[1].zfill(episode_padding)}.{words_list[-1]}")
		elif(webseries_num == 2): 
		    os.rename(f"corrected_srt\\{dictn[webseries_num]}\\{file_name}", f"corrected_srt\\{dictn[webseries_num]}\\{dictn[webseries_num]} - Season {nums_list[0].zfill(season_padding)} Episode {nums_list[1].zfill(episode_padding)} -{ep_name[-1]}.{words_list[-1]}")
		elif(webseries_num == 3): 
		    os.rename(f"corrected_srt\\{dictn[webseries_num]}\\{file_name}", f"corrected_srt\\{dictn[webseries_num]}\\{dictn[webseries_num]} - Season {nums_list[0].zfill(season_padding)} Episode {nums_list[1].zfill(episode_padding)} -{ep_name[-1]}.{words_list[-1]}")

regex_renamer()