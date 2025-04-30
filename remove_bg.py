# Importing Required Modules 
from rembg import remove 
from PIL import Image 

# Store path of the image in the variable input_path 
input_path = r'C:\Users\mabba\Downloads\Telegram Desktop\DSC_2724.jpg' 

# Store path of the output image in the variable output_path 
output_path = r"C:\Users\mabba\Downloads\Telegram Desktop\DSC_2724output.jpg"

# Processing the image 
input = Image.open(input_path) 

# Removing the background from the given Image 
output = remove(input) 

# Convert the image to RGB mode
output = output.convert("RGB")

#Saving the image in the given path 
output.save(output_path) 
