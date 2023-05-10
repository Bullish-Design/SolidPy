# In Between step for NoteGraphDB. This project's goal is to create the interface functionality to various aspects of the job into Dendron.
#   That functionality will later be extended to the graph database model, with the benefits and drawbacks that enatils.

# Steps:
#  3. Solidworks:
#     Goal: Incremental functionality integration of Solidworks with Dendron, and the ability to create+update Solidworks files from Dendron.
#     a. Solidworks Create file from Excel (VBA macro)
#       1. Generate nested folder+csv structure with excel file at root
#       2. Generate Solidworks file from csv
#       3. Generate Solidworks assembly from csv
#       4. Populate Solidworks file with features from csv
#       5. Update CSV with New/Changed Solidworks file information
#     b. Excel flexible import from CSV (VBA macro)
#     c. CSV update from Dendron (Python Script)


# VBA Macro to create new Solidworks file from Excel file


# VBA Macro to get/create/update Solidworks file properties from/to Excel file


# VBA Macro to get/create/update Solidworks file features from/to Excel file


# ----------------------------------------- Code Starts Here -----------------------------------------

# ----------------------------------------- Imports -----------------------------------------
from dataclasses import dataclass, field

# ----------------------------------------- Classes -----------------------------------------

# Data Read/Write Class
@dataclass
class swDataItem:
    table: str
    rowID: str
    column: str

    def read(self):

        pass

    def write(self):

        pass


# SW Class:
@dataclass
class Solidworks:
    id: int


# File Class:

# Part File Class:
@dataclass
class Part:
    id: int
    fileName: str
    filePath: str
    fileProps: list = field(init=False)  # List[FileProperty] - FileProperties Class?
    configurations: list = field(
        init=False
    )  # List[Configurations] - Configurations Class?


@dataclass
class FileProperty:
    id: int
    name: str
    propName: str
    propValue: str


@dataclass
class Configurations:
    id: int
    name: str
    features: list = field(init=False)  # List[Features] - Features Class?


@dataclass
class Features:
    id: int
    name: str
    featureType: str
    featSketch: list = field(init=False)  # List[Sketches] - Sketches Class?
    featProps: list = field(
        init=False
    )  # list[FeatureProperties] - FeatureProperties Class?


@dataclass
class FeatureProperties:
    id: int
    name: str
    propType: str
    propValue: str


@dataclass
class Sketch:
    id: int
    name: str
    sketchType: str
    sketchProps: list = field(init=False)


# Assembly File Class:
@dataclass
class Assembly:
    id: int
    fileName: str
    filePath: str
    asmParts: list = field(init=False)  # List[Part] - Parts Class?
    matesList: list = field(init=False)  # List[Mate] - Mates Class?
    configurations: list = field(init=False)


@dataclass
class Mate:
    id: int
    name: str
    mateType: str
    mateProps: list = field(init=False)  # List[MateProperty] - MateProperties Class?


@dataclass
class MateProperty:
    id: int
    name: str
    matePropType: str
    matePropValue: str


# Drawing File Class:

# Document Properties Class:

# ----------------------------------------- Part Features Class Tree -----------------------------------------
# Part Features Class:

# Plane Feature Class:

# Sketch Feature Class:

# Extrude Feature Class:

# Cut Feature Class:

# Revolve Feature Class:

# Line Feature Class:

# Circle Feature Class:

# Arc Feature Class:

# Common Shapes Feature Class:

# ----------------------------------------- Master Skeleton Class Tree -----------------------------------------


# ----------------------------------------- Functions -----------------------------------------

# Sync updates to Solidworks file from Dendron
def syncSWFile():
    pass


def getSWfileID():
    # Get Solidworks file ID from Active File
    if checkSWfileID() == True:
        # Get Solidworks file ID from current file
        pass
    #
    pass


def checkSWfileID() -> bool:
    # Check if Solidworks file ID exists in current file

    pass


def createSWfileID():
    # Create Solidworks file ID in current file:
    #    - generateID()
    #    - createSW_DendronFile()
    #    - updateSW_DendronFile()

    pass


def generateID():  # -> swID:
    # Generate random Solidworks file ID
    # Use structural pattern matching to use a regex-match ID typing system
    # Check to make sure ID isn't used already somehow. Create DB to store IDs and associated info.
    #   - ID Prefix (Table name in SQLdb - SLD, NSU, etc.)
    #   - ID Type (ID prefix specific tyoe - SW prefix has: PRT, ASM, DRW, PRP, etc.)
    #   - ID Number (ID prefix specific number)

    pass


def createSW_DendronFile():

    pass


def updateSW_DendronFile():

    pass
