# drawio decoder
from select import select
from tkinter import N
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
            if key.lower()=='cmdb':
                setattr(self, 'cmdb', attributes[key])

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
        self.source = None
        self.target = None
        self.source_point = None
        self.target_point = None
        super().__init__(attributes)

    def print(self):
        if self.source_point is not None:
            print(f'source point: {self.source_point[0]}, {self.source_point[1]}')
        if self.target_point is not None:
            print(f'target point: {self.target_point[0]}, {self.target_point[1]}')
        return super().print()

class Element (Object):
    def __init__(self,attributes):
        self.left_top = None
        self.right_bottom = None
        self.parent_id = None
        super().__init__(attributes)

    def is_element_inside(self,parent_element):
        if parent_element.left_top is None or parent_element.right_bottom is None:
            return False
        if self.left_top is None or self.right_bottom is None:
            return False
        if self.left_top[0] >= parent_element.left_top[0] and self.left_top[1] >= parent_element.left_top[1] and self.right_bottom[0] <= parent_element.right_bottom[0] and self.right_bottom[1] <= parent_element.right_bottom[1]:
            return True
        return False

# function that export to xls
def export_to_xls(outputfile,components,relations):
    workbook = xlsxwriter.Workbook(outputfile)

    worksheet_components = workbook.add_worksheet("Components")
    component_attribute_map = {}
    i = 0
    for comp in components.values():
        for key in comp.__dict__.keys():
            if key not in component_attribute_map:
                if key not in ['left_top', 'right_bottom']:
                    component_attribute_map[key] = i
                    worksheet_components.write(0,i,key)
                    i = i +1

    j = 1
    for component in components.values():
        for key in component.__dict__.keys():
                if key in component_attribute_map:
                    worksheet_components.write(j,component_attribute_map[key],component.__dict__[key])
        j = j + 1


    worksheet_relations = workbook.add_worksheet("Relations")
    relation_attribute_map = {}
    i = 0
    for rel in relations:
        for key in rel.__dict__.keys():
            if key not in relation_attribute_map:
                if key not in ['source_point','target_point']:
                    relation_attribute_map[key] = i
                    worksheet_relations.write(0,i,key)
                    i = i +1

    j = 1
    for rel in relations:
        for key in rel.__dict__.keys():
            if key in relation_attribute_map:
                worksheet_relations.write(j,relation_attribute_map[key],rel.__dict__[key])
        j = j + 1

    workbook.close()

# function that export to Structurizr DSL

symbols = [u"абвгдеёжзийклмнопрстуфхцчшщъыьэюяАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ abvgdeejzijklmnoprstufhzcss_y_euaABVGDEEJZIJKLMNOPRSTUFHZCSS_Y_EUA_",
           u"abvgdeejzijklmnoprstufhzcss_y_euaABVGDEEJZIJKLMNOPRSTUFHZCSS_Y_EUA_abvgdeejzijklmnoprstufhzcss_y_euaABVGDEEJZIJKLMNOPRSTUFHZCSS_Y_EUA_"]

def create_var_name(name,dubles,deep):
    prefix = 'var'
    if deep == 1 : prefix = 'system_'
    elif deep == 2: prefix = 'container_'
    elif deep == 3: prefix = 'component_'

    #name = prefix+str(len(names))

    res = ""
    src = name.lower()
    for c in src:
        for i in range(len(symbols[0])):
            if c == symbols[0][i]:
                res += symbols[1][i]
                break
    
    if res in dubles:
        new_res = res+str(len(dubles[res]))
        dubles[res].append(res)
        res = new_res
    else:
        dubles[res] = [res]

    return res

def recurse_walk(components,relations,file,component,deep,names,visible_names,dubles):
    child_count = 0

    id   = component.id
    name = component.c4Name.replace("\n"," ")
    if len(name)==0:
        name = component.c4Type.replace("\n"," ")

    if name in visible_names:
        new_name = name + '_'+str(len(visible_names[name]))
        visible_names[name].append(new_name)
        name = new_name
    else:
        visible_names[name] = list()

    var_name = create_var_name(name,dubles,deep)
    var_type = 'system'

    if 'c4Type' in component.__dict__:
        if component.c4Type != None:
            var_type = component.c4Type

    if deep == 1:
        if var_type == 'Person':
            file.write('    '+var_name+' = Person "'+name +'" {\n')
        else: 
            file.write('    '+var_name+' = softwareSystem "'+name +'" {\n')
            if hasattr(component,'cmdb'):
                file.write("        properties {\n")
                file.write(f"           cmdb {component.cmdb}\n")
                file.write("        }\n")
    elif deep == 2:
        file.write('        '+var_name+' = container "'+name +'" {\n')
    elif deep == 3:
        file.write('            '+var_name+' = component "'+name +'" {\n')

    if 'c4Description' in component.__dict__:
        if component.c4Description != None:
            for i in range(deep):
                file.write('    ')
            file.write('description "'+component.c4Description.replace("\n"," ")+'"\n')

    if 'c4Technology' in component.__dict__:
        if component.c4Technology != None:
            for i in range(deep):
                file.write('    ')
            file.write('technology "'+component.c4Technology.replace("\n"," ")+'"\n')

    for comp in components.values():
        if comp.parent_id==id:
            recurse_walk(components,relations,file,comp,deep+1,names,visible_names,dubles)
            child_count += 1

    names.append([var_name,deep,child_count,id])

    if deep == 1:
        file.write('    }\n')
    elif deep == 2:
        file.write('        }\n')
    elif deep == 3:
        file.write('            }\n')



def export_to_dsl(components,relations):
    
    with open("workspace.dsl","w") as file:
        file.write("workspace {\n")
        file.write("model {\n")
        names = list()
        visible_names = dict()

        i = 1
        dubles = dict()
        for comp in components.values():
            if 'parent_id' in comp.__dict__.keys() and comp.parent_id!=None:
                # do nothing
                i = i+1  
            else:
                recurse_walk(components,relations,file,comp,1,names,visible_names,dubles)

        elements = dict()
        for n in names:
            elements[n[3]] = n

        rel_names = dict()
        for rel in relations:
            rel_name = rel.c4Description.replace("\n"," ")
            if(len(rel_name) == 0):
                rel_name = 'Вызов'
            rel_technology = rel.c4Technology.replace("\n"," ")
            if(len(rel_technology) == 0):
                rel_technology = 'unknown'
            file.write("    "+elements[rel.source][0]+" -> "+elements[rel.target][0]+" \""+rel_name+"\" \""+rel_technology+"\"\n")

        file.write("}\n")
        file.write("views {\n")

        file.write("    systemLandscape {\n")
        file.write("        include *\n")
        file.write("        autoLayout\n")
        file.write("    }\n")

        for n in names:
            if n[2]>0 : # have childs
                if n[1] == 1:
                    file.write("    container "+n[0]+" {\n")
                    file.write("        include *\n")
                    file.write("        autoLayout\n")
                    file.write("    }\n")
                elif n[1] == 2:
                    file.write("    component "+n[0]+" {\n")
                    file.write("        include *\n")
                    file.write("        autoLayout\n")
                    file.write("    }\n")


        file.write("    themes default\n")
        file.write("    }\n")
        file.write("}\n")

# helper function to get coordinates
def get_coordinates(collection):
    coordinates = []
    coordinates.append(float(0))
    coordinates.append(float(0))
    if 'x' in collection.keys():
        coordinates[0] = float(collection['x'])
    if 'y' in collection.keys():
        coordinates[1] = float(collection['y'])
    return coordinates

# function that load from xml (.drawio)
def load_from_xml(filename,print_statistics):
    xml = open(filename).read()
    components = {}
    relations = []
    broken_relations = []

    xml_document = ET.ElementTree(ET.fromstring(xml))
    diagram_element = xml_document.find("diagram")

    if not diagram_element is None:
        if list(diagram_element): # unencoded
            root_node =list(diagram_element)[0]
            #print(ET.tostring(root_node))
        else: # encoded
            b64 = diagram_element.text
            a = base64.b64decode(b64)
            b = pako_inflate_raw(a)
            c = js_decode_uri_component(b.decode())
            #print(c)
            root_node = ET.ElementTree(ET.fromstring(c))

        for d in root_node.findall('root/object'):
            if 'c4Type' in d.attrib:
                # parse c4 relations
                if d.attrib['c4Type'] == 'Relationship':
                    mx_cell = d.find('mxCell')
                    if(mx_cell is not None):
                        have_source = False
                        have_target = False
                        source = None
                        target = None
                        if 'source' in mx_cell.attrib:
                            source = mx_cell.attrib['source']
                            have_source = True
                        if 'target' in mx_cell.attrib:
                            target = mx_cell.attrib['target']
                            have_target = True

                        if have_source and have_target: 
                            rel = Relation(source, target,d.attrib)
                            if not 'c4Description' in d.attrib:
                                rel.__setattr__('c4Description','')
                            if not 'c4Name' in d.attrib:
                                rel.__setattr__('c4Name','')
                            if not 'c4Technology' in d.attrib:
                                rel.__setattr__('c4Technology','')       
                            relations.append(rel)
                        else:
                            # case then component have no source or target
                            broken_relation = BrokenRelation(d.attrib)
                            if have_source:
                                broken_relation.source = source
                            if have_target:
                                broken_relation.target = target

                            # try to get infoermation of source and target point from relations
                            geom = mx_cell.find('mxGeometry')
                            if geom is not None:
                                points = geom.findall('mxPoint')
                                for p in points:
                                    if 'as' in p.attrib:
                                        if p.attrib['as'] == 'source' or p.attrib['as'] == 'sourcePoint':
                                            broken_relation.source_point = get_coordinates(geom.attrib)
                                        if p.attrib['as'] == 'target' or p.attrib['as'] == 'targetPoint':
                                            broken_relation.target_point = get_coordinates(geom.attrib)
                            if(not 'c4Description' in d.attrib):
                                broken_relation.__setattr__('c4Description','')
                            if(not 'c4Name' in d.attrib):
                                broken_relation.__setattr__('c4Name','')
                            if(not 'c4Technology' in d.attrib):
                                broken_relation.__setattr__('c4Technology','')
                            broken_relations.append(broken_relation)
                else:
                    # parse c4 components
                    comp = Element(d.attrib)

                    mx_cell = d.find('mxCell')
                    if(mx_cell is not None):
                        geom = mx_cell.find('mxGeometry')
                        if geom is not None:
                            comp.left_top = get_coordinates(geom.attrib)
                            comp.right_bottom = [comp.left_top[0] + float(geom.attrib['width']),comp.left_top[1] + float(geom.attrib['height'])]
                    components[comp.id] = comp

        # parse labels and edges for non-c4 relations            
        labels = {}
        for d in root_node.findall('root/mxCell'):   
            if 'style' in d.attrib:
                # parse edge
                if d.attrib['style'].find('edgeStyle=') != -1:
                    broken_relation = BrokenRelation({})
                    broken_relation.id = d.attrib['id']
                    if 'source' in d.attrib:
                        broken_relation.source = d.attrib['source']
                    if 'target' in d.attrib:
                        broken_relation.target = d.attrib['target']
                    broken_relation.__setattr__('c4Name','')
                    broken_relation.__setattr__('c4Type','Relationship')
                    broken_relation.__setattr__('c4Technology','')
                    broken_relation.__setattr__('c4Description','')
                    broken_relations.append(broken_relation)
            
            # parse label
            if 'style' in d.attrib:
                if d.attrib['style'].find('edgeLabel') != -1:
                    if( 'parent' in d.attrib) and ('value' in d.attrib):
                            labels[d.attrib['parent']] = d.attrib['value']

        # parse technology from non c4-relations labels
        for label in labels.keys():
            parents = [x for x in broken_relations if x.id == label]
            if len(parents) > 0:   
                parents[0].c4Description = labels[label]
                m = re.search(r'\[(.*)\]', labels[label])
                if m:
                    parents[0].c4Technology = m.group(1)

    if print_statistics==True:
        print('Number of components: ' + str(len(components)))
        print('Number of relations: ' + str(len(relations)))
        print('Number of broken relations: ' + str(len(broken_relations)))


    return components, relations ,broken_relations

# remove relationship that links to component that not in component list
def fix_missing_relations(components,relations):
    result_relations = []
    for rel in relations:
        if rel.source not in components.keys():
            rel.source = None
        if rel.target not in components.keys():
            rel.target = None

        if rel.source is not None and rel.target is not None:
            result_relations.append(rel)
    return result_relations

# fix broken relations
def fix_broken_relations(components,relations,broken_relations):
    i = 0
    for broken_relation in broken_relations:
        if broken_relation.source is None and broken_relation.source_point is not None:
            candidats = {}
            for comp in components.values():
                if comp.left_top[0] <= broken_relation.source_point[0] <= comp.right_bottom[0] and comp.left_top[1] <= broken_relation.source_point[1] <= comp.right_bottom[1]:                                       
                    candidats[(comp.right_bottom[0]-comp.left_top[0])*(comp.right_bottom[1]-comp.left_top[1])] = comp.id;                  
                    
            if len(candidats) > 0:
                broken_relation.source = candidats[min(candidats.keys())]

        if broken_relation.target is None and broken_relation.target_point is not None:
            candidats = {}
            for comp in components.values():
                if comp.left_top[0] <= broken_relation.target_point[0] <= comp.right_bottom[0] and comp.left_top[1] <= broken_relation.target_point[1] <= comp.right_bottom[1]:
                    candidats[(comp.right_bottom[0]-comp.left_top[0])*(comp.right_bottom[1]-comp.left_top[1])] = comp.id;  
                    
            if len(candidats) > 0:
                broken_relation.target = candidats[min(candidats.keys())]

        if broken_relation.source is not None and broken_relation.target is not None:
            i = i + 1
            #print(broken_relation.__dict__)
            relations.append(Relation(broken_relation.source,broken_relation.target,broken_relation.__dict__))
            
    return relations

# function that print broken relations
def print_broken_relations(broken_relations,i):
    for br in broken_relations:
        print(f'{i}. Связь {br.id} "{br.c4Name}" не имеет начала или конца "{br.c4Description}"')
        if br.source is not None:
            print(f'Начало: {br.source}')
        if br.target is not None:
            print(f'Конец: {br.target}')
        
        if br.source_point is not None:
            print(f'Начало: {br.source_point}')
        if br.target_point is not None:
            print(f'Конец: {br.target_point}')
        i = i+1
    return i

# function that check relations
def check_relations(components, relations,i,check_data):
    def component_name(component):
        if len(component.c4Name)!=0:
            return component.c4Name.replace('\n',' ')
        else:
            return component.c4Type+":"+component.c4Description.replace('\n',' ')

    def relation_name(relation):
        if(len(relation.c4Description.rstrip())>0):
            # return relation.c4Description with replaced newlines
            return relation.c4Description.replace('\n',' ')
        else:
            return ''

    for rel in relations:
        if rel.source not in components:
            print(f'Для связи "{relation_name(rel)}" отсутствует стартовый компонент')
        if rel.target not in components:
            print(f'Для связи "{relation_name(rel)}" отсутствует конечный компонент')
        if 'c4Technology' in rel.__dict__:
            if rel.c4Technology=='' and components[rel.source].c4Type != 'Person' and components[rel.target].c4Type != 'Person':
                print(f'{i}. Для связи "{relation_name(rel)}" между "{component_name(components[rel.source])}" и "{component_name(components[rel.target])}" не указана технология')
                i = i + 1
        if 'c4Description' in rel.__dict__ and check_data:
            m = re.search(r'\((.*)\)', rel.c4Description)
            if m is None:
                if components[rel.source].c4Type != 'Person' and components[rel.target].c4Type != 'Person':
                    print(f'{i}. Для связи "{relation_name(rel)}" между "{component_name(components[rel.source])}" и "{component_name(components[rel.target])}" не указаны входные данные')
                    i = i + 1
            m = re.search(r'\):(.*)', rel.c4Description)
            if m is None:
                if components[rel.source].c4Type != 'Person' and components[rel.target].c4Type != 'Person':
                    print(f'{i}. Для связи "{relation_name(rel)}" между "{component_name(components[rel.source])}" и "{component_name(components[rel.target])}" не указаны возвращаемые данные')
                    i = i + 1
    return i

# function that fills parent id
def fill_parent_id(components):
    result = {}
    for comp in components.values():
        for parent in components.values():
            if comp != parent:
                if comp.is_element_inside(parent):
                    comp.parent_id = parent.id
        result[comp.id] = comp
    return result

# function that checks inpound and outbound relations
# if comonent is habe parent component, than parent must have inbound or outbound relation
def check_inbound_outbound_relations(comp,components,relations):
    if not [x for x in relations if x.target == comp.id or x.source == comp.id]:
        if comp.parent_id is not None:
            return check_inbound_outbound_relations(components[comp.parent_id],components,relations)
        else:
            return False
    else:
        return True


# function that checks components
def check_components(components, relations, i):
    for comp in components.values():
        if 'c4Description' not in comp.__dict__:
            if comp.c4Type != 'SystemScopeBoundary' and comp.c4Type != 'ContainerScopeBoundary' and comp.c4Type != 'Person':
                print(f'{i}. {comp.c4Type} "{comp.c4Name}" не указано описание')
                i = i + 1
        if 'c4Technology' not in comp.__dict__:
            if(comp.c4Type != 'Software System') and (comp.c4Type != 'Person') and (comp.c4Type != 'SystemScopeBoundary') and (comp.c4Type != 'ContainerScopeBoundary'):
                print(f'{i}. {comp.c4Type} "{comp.c4Name}" не указана технология')
                i = i + 1
        
        if comp.c4Type != 'SystemScopeBoundary' and comp.c4Type != 'Person' and comp.c4Type != 'ContainerScopeBoundary':
            if check_inbound_outbound_relations(comp,components,relations) is False:
                print(f'{i}. {comp.c4Type} "{comp.c4Name}" не имеет входящих и исходящих связей')
                i = i + 1
    return i

# main function
def main(argv):
    # parse args
    inputfile = ''
    outputfile = ''
    check_data = False
    print_statistics = False

    helpstring = 'drawio_parser.py -i <inputfile> -o <outputfile> -d -s'
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
        elif opt in ("-o", "--ofile"):
            outputfile = arg
        elif opt == '-d':
            check_data = True
        elif opt == '-s':
            print_statistics = True

    if len(inputfile) == 0:
        print (helpstring)
        sys.exit()


    # load from xml (.drawio)
    components, relations , broken_relations = load_from_xml(inputfile,print_statistics)

    # fill parent relations
    components = fill_parent_id(components)

    # fix broken relations
    relations = fix_broken_relations(components, relations, broken_relations)


    relations = fix_missing_relations(components, relations)
    # make checks
    i = 1
    #i = print_broken_relations(broken_relations,i)
    i = check_relations(components, relations, i,check_data)
    i = check_components(components, relations, i)


    # export to xls
    if len(outputfile) != 0:
        export_to_xls(outputfile,components,relations)
    
    export_to_dsl(components,relations)

if __name__ == "__main__":
   main(sys.argv[1:])
