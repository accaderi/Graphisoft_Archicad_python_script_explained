# import archicad connection (required)
from archicad import ACConnection

# Establish the connection with the Archicad software, Archicad must be open and the pln file must be open too.
#
# We can use the acu.OpenFile() utility but we need an established connection first so we need to open Archicad and a new plan
# as a minimum because all utilities use the connection.
conn = ACConnection.connect()
# assert that the connection is alive
assert conn

# Create shorts of the commands, types and utilities.
acc = conn.commands
act = conn.types
acu = conn.utilities

# original comment -> ################################ CONFIGURATION #################################
# Getting the property Id (guid) of the user defined 'Zone Overall' property
# from the Zones group in order to use this to identify
# the property exactly when we communicate with the API
# (this is a unique identifier like our social security number).
propertyId = acu.GetUserDefinedPropertyId("ZONES", "Zone Overall")
# We collect all the 'Zone' element into the 'elements' list
# using the 'Zone' type with the GetElementsByType command.
# The elements list contains the 'guids' of all the the zones in the project.
elements = acc.GetElementsByType('Zone')

# With this function we generate a string property value
# since the user defined 'Zone Overall' is a string type property.
# Takes as argument the width and height
# returns a string: width x height or height x width
# taking the highest of these two first.
def GeneratePropertyValueString(width: float, height: float) -> str:
    # original comment -> # show highest value first - office preference.
    if width > height:
        return f"{width:.2f} x {height:.2f}"
    else:
        return f"{height:.2f} x {width:.2f}"
# original comment -> ################################################################################

# This function prepares NormalStringPropertyValue type from the string
# generated by the function 'GeneratePropertyValueString'.
# The function of GeneratePropertyValueString could be combined with this
# since these two functions are working always together.
def generatePropertyValue(width: float, height: float) -> act.NormalStringPropertyValue:
    return act.NormalStringPropertyValue(GeneratePropertyValueString(width, height))


# Getting all 2d bounding boxes of all the zones.
# The 2d bounding box contains the x, y minimum and maximum values of the box
# can be drawn around the element containging the whole element!
# Returns a list.
# original comment -> # collect all the data
boundingBoxes = acc.Get2DBoundingBoxes(elements)

# Dictionary of each element and its bounding box follows.
# Calling zip method on the elements list and the bounding box list.
# original comment -> # bind bounding boxes to element ids
elementBoundingBoxDict = dict(zip(elements, boundingBoxes))

# Initialise the elemntPropertyValues list to collect all the final elementproperty values.
# original comment -> # calculated the widths and heights
elemPropertyValues = []
# This for loop is filling up the above created list calculating the width and height.
for key,value in elementBoundingBoxDict.items():
        # width = xMax - xMin
        width = abs(value.boundingBox2D.xMax - value.boundingBox2D.xMin)
        # height = yMax - yMin
        height = abs(value.boundingBox2D.yMax - value.boundingBox2D.yMin)
        
        # Generate the property value calling the generatePropertyValue function. 
        newPropertyValue = generatePropertyValue(width, height)
        # Generate and append the element property values.
        # ElementPropertyValue type takes arguments:
        # elementId, propertyId, property value
        elemPropertyValues.append(act.ElementPropertyValue(key.elementId, propertyId, newPropertyValue))
 
# original comment -> # set the new property values
# Argument: elemPropertyValues list. (It takes only list even if we have one element)
acc.SetPropertyValuesOfElements(elemPropertyValues)