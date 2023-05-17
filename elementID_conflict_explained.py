# Import archicad connection (required).
from archicad import ACConnection

# Establish the connection with the Archicad software, Archicad must be open and the pln file must be open too.
#
# We can use the acu.OpenFile() utility but we need an established connection first so we need to open Archicad and a new plan
# as a minimum because all utilities use the connection.
conn = ACConnection.connect()
assert conn

# Create shorts of the commands, types and utilities.
acc = conn.commands
act = conn.types
acu = conn.utilities

# original comment -> ################################ CONFIGURATION #################################
# Get all elements from the project.
elements = acc.GetAllElements()
# Define messages.
messageWhenNoConflictFound = "There is no elementID conflict."
conflictMessageParts = ["[Conflict]", "elements have", "as element ID:\n"]

# This function createss the constructed message for the conflicts.
def GetConflictMessage(elementIDPropertyValue, elementIds):
    return f"{conflictMessageParts[0]} {len(elementIds)} {conflictMessageParts[1]} '{elementIDPropertyValue}' {conflictMessageParts[2]}{sorted(elementIds, key=lambda id: id.guid)}"
# original comment -> ################################################################################

# Get the built in property id of 'General_ElementID' for all the elements.
elementIdPropertyId = acu.GetBuiltInPropertyId('General_ElementID')
# Get the built in property value of 'General_ElementID' for all the elements.
propertyValuesForElements = acc.GetPropertyValuesOfElements(elements, [elementIdPropertyId])

# Create a dictionary with the property value : element id, key:value pairs.
propertyValuesToElementIdsDictionary = {}
# Loop through the elements' property values.
for i in range(len(propertyValuesForElements)):
    # Take the element id of the actual element.
    elementId = elements[i].elementId
    # Take the property value of the actual element.
    propertyValue = propertyValuesForElements[i].propertyValues[0].propertyValue.value
    # If the property value is not in the dictionary create the key of the actual property value
    # and its value as an empty set.
    if propertyValue not in propertyValuesToElementIdsDictionary:
        propertyValuesToElementIdsDictionary[propertyValue] = set()
    # Add the actual elmentid to the set.
    propertyValuesToElementIdsDictionary[propertyValue].add(elementId)

# As base condition no conflict.
noConflictFound = True
# Loop through the 'propertyValuesToElementIdsDictionary' items.
for k, v in sorted(propertyValuesToElementIdsDictionary.items()):
    # If the actual property value's (key) set has more than one element
    # it means the same id is given to those elements which ids in the actual set.
    # It is a conflict. 
    if len(v) > 1:
        noConflictFound = False
        print(GetConflictMessage(k, v))
# If there was no conflict (the element id set of the actual value contains only one element).
if noConflictFound:
    print(messageWhenNoConflictFound)
