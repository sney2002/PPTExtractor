#-*- coding: utf-8 -*-
"""
Extract images from PowerPoint files


Extract images from PowerPoint files (ppt, pps, pptx) without use win32 API

required:
    OleFileIO_PL: Copyright (c) 2005-2010 by Philippe Lagadec

Usage

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
"""
# Copyright (c) 2010 Jhonathan Salguero Villa (http://github.com/sney2002)
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

import OleFileIO_PL as OleFile
import zipfile
import struct
import os


DEBUG = False
CWD = '.'
CHUNK = 1024 * 64

# MS-ODRAW spec
formats = {
    # 2.2.24
    (0xF01A, 0x3D40): (50, ".emf"),
    (0xF01A, 0x3D50): (66, ".emf"),
    # 2.2.25
    (0xF01B, 0x2160): (50, ".wmf"),
    (0xF01B, 0x2170): (66, ".wmf"),
    # 2.2.26
    (0xF01C, 0x5420): (50, ".pict"),
    (0xF01C, 0x5430): (50, ".pict"),
    # 2.2.27
    (0xF01D, 0x46A0): (17, ".jpeg"),
    (0xF01D, 0x6E20): (17, ".jpeg"),
    (0xF01D, 0x46B0): (33, ".jpeg"),
    (0xF01D, 0x6E30): (33, ".jpeg"),
    # 2.2.28
    (0xF01E, 0x6E00): (17, ".png"),
    (0xF01E, 0x6E10): (33, ".png"),
    # 2.2.29
    (0xF01F, 0x7A80): (17, ".dib"),
    (0xF01F, 0x7A90): (33, ".dib"),
    # 2.2.30
    (0xF029, 0x6E40): (17, ".tiff"),
    (0xF029, 0x6E50): (33, ".tiff")
}

class InvalidFormat(Exception):
    pass


class PowerPointFormat(object):
    def __init__(self, filename):
        """
        filename:   archivo a abrir
        """
        self._files = {}
        
        # nombre base de imágenes
        self.basename = os.path.splitext(os.path.basename(filename))[0]
        
        self._process(filename)

    def extract(self, name, newname="", path=CWD):
        """
        Extrae imagen en directorio especificado.
        """
        self._extract(name, newname=newname, path=path)
        
    def extractall(self, path=CWD):
        """
        Extrae todas las imágenes en directorio especificado.
        """
        for img in self._files:
            self.extract(img, path=path)
            
    def namelist(self):
        """
        Retorna lista de imágenes contenidas en el archivo.
        """
        return self._files.keys()
        
    def __len__(self):
        return len(self._files)
        
    def __str__(self):
        return "<PowerPoint file with %s images>" % len(self)
    
    __repr__ = __str__


# TODO: Extraer otros tipos de archivo (wav, avi...)
class PPT(PowerPointFormat):
    """
    Extrae imágenes de archivos PowerPoint binarios (ppt, pps).
    """
    headerlen = struct.calcsize('<HHL')

    @classmethod
    def is_valid_format(cls, filename):
        return OleFile.isOleFile(filename)
    
    def _process(self, filename):
        """
        Busca imágenes dentro de stream y guarda referencia a su ubicación.
        """
        olefile = OleFile.OleFileIO(filename)
        
        # Al igual que en pptx esto no es un error
        if not olefile.exists("Pictures"):
            return
            #raise IOError("Pictures stream not found")
        
        self.__stream = olefile.openstream("Pictures")

        stream = self.__stream
        offset = 0
        # cantidad de imágenes encontradas
        n = 1

        while True:
            header = stream.read(self.headerlen)
            offset += self.headerlen

            if not header: break
            
            # cabecera
            recInstance, recType, recLen = struct.unpack_from("<HHL", header)

            # mover a siguiente cabecera
            stream.seek(recLen, 1)

            if DEBUG:
                print "%X %X %sb" % (recType, recInstance, recLen)
            
            extrabytes, ext = formats.get((recType, recInstance))
            
            # Eliminar bytes extra
            recLen -= extrabytes
            offset += extrabytes
            
            # Nombre de Imagen
            filename = "{0}{1}{2}".format(self.basename, n, ext)
            
            self._files[filename] = (offset, recLen)
            offset += recLen
            
            n += 1

    def _extract(self, name, newname="", path=CWD):
        """
        Extrae imagen en el directorio actual (path).
        """
        filename = newname or name
        
        if not name in self._files:
            raise IOError("No such file")
        
        offset, size = self._files[name]
        
        # dirección de destino completa
        filepath = os.path.join(path, filename)
        
        total = 0
        
        self.__stream.seek(offset, 0)

        with open(filepath, "wb") as output:
            while (total + CHUNK) < size:
                data = self.__stream.read(CHUNK)
                
                if not data: break
                
                output.write(data)
                total += len(data)
                
            if total < size:
                data = self.__stream.read(size - total)
                output.write(data)

class PPTX(PowerPointFormat):
    """
    Extrae imágenes de archivos PowerPoint +2007
    """
    @classmethod
    def is_valid_format(cls, filename):
        return zipfile.is_zipfile(filename)
    
    def _process(self, filename):
        """
        Busca imágenes dentro de archivo zip y guarda referencia a su ubicación.
        """
        self.__zipfile = zipfile.ZipFile(filename)
        
        n = 1

        for file in self.__zipfile.namelist():
            path, name = os.path.split(file)
            name, ext = os.path.splitext(name)
            
            # los archivos multimedia se guardan en ppt/media
            if path == "ppt/media":
                filename = "{0}{1}{2}".format(self.basename, n, ext)
                
                # guardar path de archivo dentro del zip
                self._files[filename] = file
                
                n += 1
                
    def _extract(self, name, newname="", path=CWD):
        """
        Extrae imagen en el directorio actual (path).
        """
        filename = newname or name
        
        if not name in self._files:
            raise IOError("No such file")
        
        # dirección de destino completa
        filepath = os.path.join(path, filename)
        
        total = 0
        
        # extraer archivo
        file = self.__zipfile.open(self._files[name])
        
        with open(filepath, "wb") as output:
            while True:
                data = file.read(CHUNK)
                
                if not data: break
                
                output.write(data)
                total += len(data)


def PPTExtractor(filename):
    # Identificar tipo de archivo (pps, ppt, pptx) e instanciar clase adecuada
    for cls in PowerPointFormat.__subclasses__():
        if cls.is_valid_format(filename):
            return cls(filename)
    raise InvalidFormat("{0} is not a PowerPoint file".format(filename))


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1:
        PPTExtractor(sys.argv[1]).extractall()
    else:
        print "Uso: %s PowerPointFile" % __file__
    
