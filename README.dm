#PPTExtractor

Extract images from PowerPoint files (ppt, pps, pptx) without use win32 API

required:
    OleFileIO_PL: Copyright (c) 2005-2010 by Philippe Lagadec

##Usage

By default images are saved in current directory:

    ppt = PPTExtractor("some/PowerPointFile")

    # found images
    len(ppt)

    # image list
    images = ppt.namelist()

    # extract image
    ppt.extract(images[0])
    
    # save image with different name
    ppt.extract(images[0], "nuevo-nombre.png")

    # extract all images
    ppt.extractall()
    
Save images in a diferent directory:
    
    ppt.extract("image.png", path="/another/directory")
    ppt.extractall(path="/another/directory")
