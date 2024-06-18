from openpyxl import load_workbook
import lxml.etree as etree
"""
XML Functions
"""
class XMLFunctions :
    @staticmethod
    def create_Item(child_data) :
        # Create an xml item, with a specific data
        # and then, return the root item
        pass

    @staticmethod
    def createXMLTree(root, root_xml, elements) :
        # create an xml tree, using a stack, LIFO : Last In First Out ==> Stack
        # return then, root_xml
        pass
    
    @staticmethod
    def create_Items_List(elements) :
        elements_xml = []
        # loop the elements and Convert the elements element to xml items .
        # return elements_xml
        pass

    @staticmethod
    def SaveXMLFile(root, output_path) :
        etree.ElementTree(root).write(output_path, pretty_print=True, encoding="utf-8")
        XMLFunctions.add_xml_declaration(output_path)
    
    @staticmethod
    def add_xml_declaration(xml_file):
    
        with open(xml_file, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        # Add the XML declaration if it doesn't exist
        if not xml_content.startswith('<?xml'):
            xml_content = '<?xml version="1.0" encoding="UTF-8" standalone="no" ?>\n' + xml_content

            # Write the modified content back to the file
            with open(xml_file, 'w', encoding='utf-8') as file:
                file.write(xml_content)

    @staticmethod
    def createNewXMLFile(root_TreeNode, elements, output_path) :   
        # Create the xml tree, by calling createXMLTree function
        
        # Then,save the xml file by calling the SaveXMLFile
        pass


def excelSheet_modulation(sheet) :
    # A loop to Delete empty rows

    # A loop to Delete empty columns

    # return the new sheet
    pass

class ExcelElementsClass :
    def __init__(self, id, name, description, niveau) :
        self.id = id
        self.name = name
        self.description = description
        self.niveau = niveau
    @staticmethod
    def getAllRowsFromExcel(sheet):
        # convert each row into a list and put it in elements, and return it .
        elements = []
        pass

class TreeNode:
    def __init__(self, data):
        self.data = data
        self.children = []

    def add_child(self, child_node):
        self.children.append(child_node)

def print_tree(node, level=0, prefix="Root"):
    # An optional function for showing the tree of TreeNode elements .
    pass


def createTree(root, elements) :
    ########
    ## RQ ##
    ## The Excel File Must Contain this headers in this order (id, name, description, niveau)
    ########
    # create a tree, using a stack, LIFO : Last In First Out ==> Stack
    pass

def createXMLFile(input_path, output_path) :
    # Create a Tree
    root = TreeNode(ExcelElementsClass('0', "Persons",None, 0))

    excelFile = load_workbook(input_path)
    workSheet = excelFile.active
    # OR :  workSheet = excelFile["Sheet1"]

    # To delete the extra empty rows and cols
    workSheet = excelSheet_modulation(workSheet)

    # contain all the excel rows with distinct value according to ('id', 'name', 'description', 'Niveau')
    elements = ExcelElementsClass.getAllRowsFromExcel(workSheet)
    
    createTree(root, elements)
    
    # call createNewXMLFile, using the Tree Root, as an XML element, and then save it as an xml file .
    XMLFunctions.createNewXMLFile(root, elements, output_path)
    
    
    # Print the tree starting from the root
    ## print_tree(root)

    # excelFile.save(output_path)
    # open_file(input_path)




