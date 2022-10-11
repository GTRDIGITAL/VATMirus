import qrcode
import png
from PIL import Image
# import _im
# features=qrcode.QRCode(version=1,box_side=40,border=3)
features=qrcode.make('https://www.youtube.com/watch?v=LoXzTW72jOw')

# link=input('https://www.youtube.com/watch?v=LoXzTW72jOw')

# generatepoza=features.make_image(fill_color="black",back_color="white")
features.save("C:/Users/Bogdan.Constantinesc/Documents/One vat app/D300 to XML Final CI/D300 to XML 2/D300 to XML/pentru cristi.png")
# im1=qr_code.png("QRCode.png",scale=5)
# im1=Image.open("QRCode.png")
# im1.save("C:/Users/Bogdan.Constantinesc/Documents/One vat app/D300 to XML Final CI/D300 to XML 2/D300 to XML/pentru cristi.png")