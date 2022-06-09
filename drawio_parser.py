# drawio decoder
from select import select
import xml.etree.ElementTree as ET
import re
import base64
import zlib
from urllib.parse import quote, unquote

# args
import sys, getopt

# xls
import xlsxwriter

def js_encode_uri_component(data):
    return quote(data, safe='~()*!.\'')


def js_decode_uri_component(data):
    return unquote(data)


def js_string_to_byte(data):
    return bytes(data, 'iso-8859-1')


def js_bytes_to_string(data):
    return data.decode('iso-8859-1')


def js_btoa(data):
    return base64.b64encode(data)


def js_atob(data):
    return base64.b64decode(data)

def pako_inflate_raw(data):
    decompress = zlib.decompressobj(-15)
    decompressed_data = decompress.decompress(data)
    decompressed_data += decompress.flush()
    return decompressed_data

# diagram elements
class Object:
    def __init__(self,attributes):
        self.id = None
        setattr(self, 'c4Name', '')
        for key in attributes.keys():
            if key == 'id':
                self.id = attributes[key]
            if key.startswith('c4'):
                setattr(self, key, attributes[key])

    def print(self):
        for key in self.__dict__.keys():
            print(key, ':', getattr(self, key))

class Relation (Object):
    def __init__(self,source,target, attributes):
        super().__init__(attributes)
        self.source = source
        self.target = target

    def print(self):
        return super().print()

class BrokenRelation (Object):
    def __init__(self, attributes):
        super().__init__(attributes)

    def print(self):
        return super().print()

class Element (Object):
    def __init__(self,attributes):
        super().__init__(attributes)

# function that export to xls
def export_to_xls(outputfile,components,relations):
    workbook = xlsxwriter.Workbook(outputfile)

    worksheet_components = workbook.add_worksheet("Components")
    component_attribute_map = {}
    for comp in components.values():
        for key in comp.__dict__.keys():
            if key not in component_attribute_map:
                component_attribute_map[key] = 0
    i = 0
    for key in component_attribute_map.keys():
        component_attribute_map[key] = i
        worksheet_components.write(0,i,key)
        i = i +1

    j = 1
    for component in components.values():
        for key in component.__dict__.keys():
                worksheet_components.write(j,component_attribute_map[key],component.__dict__[key])
        j = j + 1

    relation_attribute_map = {}
    worksheet_relations = workbook.add_worksheet("Relations")
    for rel in relations:
        for key in rel.__dict__.keys():
            if key not in relation_attribute_map:
                relation_attribute_map[key] = 0
    i = 0
    for key in relation_attribute_map.keys():
        relation_attribute_map[key] = i
        worksheet_relations.write(0,i,key)
        i = i +1

    j = 1
    for rel in relations:
        for key in rel.__dict__.keys():
                worksheet_relations.write(j,relation_attribute_map[key],rel.__dict__[key])
        j = j + 1

    workbook.close()

# function that load from xml (.drawio)
def load_from_xml(filename):
    xml = open(filename).read()
    components = {}
    relations = []
    broken_relations = []

    xml_document = ET.ElementTree(ET.fromstring(xml))
    diagram_element = xml_document.find("diagram")

    if not diagram_element is None:
        if list(diagram_element): # unencoded
            root_node =list(diagram_element)[0]
        else: # encoded
            b64 = diagram_element.text
            a = base64.b64decode(b64)
            b = pako_inflate_raw(a)
            c = js_decode_uri_component(b.decode())
            root_node = ET.ElementTree(ET.fromstring(c))

        for d in root_node.iter('object'):
            if 'c4Type' in d.attrib:
                if d.attrib['c4Type'] == 'Relationship':
                    mx_cell = d.find('mxCell')
                    if(mx_cell is not None):
                        if 'source' in mx_cell.attrib and 'target' in mx_cell.attrib:
                            source = mx_cell.attrib['source']
                            target = mx_cell.attrib['target']
                            relations.append(Relation(source, target,d.attrib))
                        else:
                            broken_relations.append(BrokenRelation(d.attrib))
                    else:
                        broken_relations.append(BrokenRelation(d.attrib))
                else:
                    comp = Element(d.attrib)
                    components[comp.id] = comp
                
    print(f"Components:{len(components)}")                
    print(f"Relations: {len(relations)}")
    print(f"Broken Relations: {len(broken_relations)}")

    return components, relations ,broken_relations

# function that print broken relations
def print_broken_relations(broken_relations,i):
    for br in broken_relations:
        print(f'{i}. Связь не имеет начала или конца "{br.c4Description}"')
        i = i+1
    return i

# function that check relations
def check_relations(components, relations,i):
    for rel in relations:
        if rel.source not in components:
            print(f"Для связи {rel.c4Description} отсутствует стартовый компонент")
        if rel.target not in components:
            print(f"Для связи {rel.c4Description} отсутствует конечный компонент")
        if 'c4Technology' in rel.__dict__:
            if rel.c4Technology=='' and components[rel.source].c4Type != 'Person' and components[rel.target].c4Type != 'Person':
                print(f'{i}. Для связи "{rel.c4Description}" между "{components[rel.source].c4Name}" и "{components[rel.target].c4Name}" не указана технология')
                i = i + 1
        if 'c4Description' in rel.__dict__:
            m = re.search(r'\((.*)\)', rel.c4Description)
            if m is None:
                if components[rel.source].c4Type != 'Person' and components[rel.target].c4Type != 'Person':
                    print(f'{i}. Для связи "{rel.c4Description}" между "{components[rel.source].c4Name}" и "{components[rel.target].c4Name}" не указаны входные данные')
                    i = i + 1
            m = re.search(r'\((.*)\):', rel.c4Description)
            if m is None:
                if components[rel.source].c4Type != 'Person' and components[rel.target].c4Type != 'Person':
                    print(f'{i}. Для связи "{rel.c4Description}" между "{components[rel.source].c4Name}" и "{components[rel.target].c4Name}" не указаны возвращаемые данные')
                    i = i + 1
    return i
    
# function that checks components
def check_components(components, relations, i):
    for comp in components.values():
        if 'c4Description' not in comp.__dict__:
            if comp.c4Type != 'SystemScopeBoundary':
                print(f'{i}. Компонент "{comp.c4Name}" не указано описание')
                i = i + 1
        if 'c4Technology' not in comp.__dict__:
            if(comp.c4Type != 'Software System') and comp.c4Type != 'Person' and comp.c4Type != 'SystemScopeBoundary':
                print(f'{i}. Компонент "{comp.c4Name}" не указана технология')
                i = i + 1
        if not [x for x in relations if x.target == comp.id]:
            if comp.c4Type != 'SystemScopeBoundary' and comp.c4Type != 'Person':
                print(f'{i}. Компонент "{comp.c4Name}" не имеет входных связей')
                i = i + 1
        if not [x for x in relations if x.source == comp.id]:
            if comp.c4Type != 'SystemScopeBoundary':
                print(f'{i}. Компонент "{comp.c4Name}" не имеет выходных связей')
                i = i + 1
    return i

def main(argv):
    # parse args
    inputfile = ''
    outputfile = ''
    helpstring = 'drawio_parser.py -i <inputfile> -o <outputfile>'
    try:
        opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
    except getopt.GetoptError:
        print (helpstring)
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print (helpstring)
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputfile = arg

    if len(inputfile) == 0 or len(outputfile) == 0:
        print (helpstring)
        sys.exit()


    # load from xml (.drawio)
    components, relations , broken_relations = load_from_xml(inputfile)

    # make checks
    i = 1
    i = print_broken_relations(broken_relations,i)
    i = check_relations(components, relations, i)
    i = check_components(components, relations, i)

    # export to xls
    export_to_xls(outputfile,components,relations)

if __name__ == "__main__":
   main(sys.argv[1:])