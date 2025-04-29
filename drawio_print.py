# drawio decoder
from select import select
from tkinter import N
import xml.etree.ElementTree as ET
from urllib.parse import quote, unquote


# args
import sys, getopt


def load_from_xml(filename):
    xml = open(filename).read()
    xml_document = ET.ElementTree(ET.fromstring(xml))
    root = xml_document.find("diagram")

    if not root is None:
        for d in root.findall('mxGraphModel/root'):
            print ('object')
            for m in d.findall('mxCell'):
                if 'value' in m.attrib:
                    print(m.attrib['value'])



# main function
def main(argv):
    # parse args
    inputfile = ''

    helpstring = 'drawio_parser.py -i <inputfile> '
    try:
        opts, args = getopt.getopt(argv,"sdhi:o:",["ifile=","ofile="])
    except getopt.GetoptError:
        print (helpstring)
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print (helpstring)
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        

    if len(inputfile) == 0:
        print (helpstring)
        sys.exit()


    load_from_xml(inputfile)



if __name__ == "__main__":
   main(sys.argv[1:])
