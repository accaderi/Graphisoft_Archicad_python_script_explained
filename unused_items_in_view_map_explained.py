# Import archicad connection (required), handle_dependencies is not necessary.
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
# These vairables can be changed based on our needs.
# Move the unused navigator items to an other folder True/False.
moveToFolder = True
# The name of the folder of the  unused views (string).
folderName = '-- UnusedViews --'
# Rename the previous unused views folder True/False.
renameFolderFromPreviousRun = True
# The name of the folder dor the previous run (string).
folderNameForPreviousRun = '-- Previous UnusedViews --'
# original comment -> ################################################################################

# This function returns True if 'sourceNavigatorItemId' does exist.
def isLinkNavigatorItem(item : act.NavigatorItem):
    return item.sourceNavigatorItemId is not None

# Getting the ##'LayoutBook' navigator item tree,
# with the 'GetNavigatorItemTree' using the 'NavigatorTreeId' type.
layoutBookTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('LayoutBook'))
# Getting all navigator items with 'sourceNavigatorItemId' in the links list.  
links = acu.FindInNavigatorItemTree(layoutBookTree.rootItem, isLinkNavigatorItem)

# In all publisher sets loop through 
for publisherSetName in acc.GetPublisherSetNames():
    # From 'PublisherSets'/actual publisher set getting the navigator item tree.
	publisherSetTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('PublisherSets', publisherSetName))
    # Adding the actual publisher set navigator items whith 'sourceNavigatorItemId' to the links list.
	links += acu.FindInNavigatorItemTree(publisherSetTree.rootItem, isLinkNavigatorItem)

# Getting all unique source links' guids to this set from the links list
# using list comprehension. 
sourcesOfLinks = set(link.sourceNavigatorItemId.guid for link in links)

# Getting the navigator item tree of the 'ViewMap'.
viewMapTree = acc.GetNavigatorItemTree(act.NavigatorTreeId('ViewMap'))
# Getting unused view tree items out of the viewMapTree if
# their name is no 'folderName' and not 'folderNameForPreviousRun'
# and not its navigator id is not in the source links ('sourcesOfLinks').
unusedViewTreeItems = acu.FindInNavigatorItemTree(viewMapTree.rootItem,
    lambda node: node.name != folderName and node.name != folderNameForPreviousRun
        # With this 'FindInNavigatorItemTree' nested 'FindInNavigatorItemTree' in our particular case
        # the code does not take the nodes and children nodes and so on into the list since we are
        # negating hence both our out, however if we would want the same as inclusive both the father and children nodes
        # would be included in the list. This would cause a bit of confusion.
        # In general the better way would be to start with the 'sourcesOfLinks' list check and then the folders check
        # to avoid the above confusion.
        # In the code below this confusion is handled with a for loop so it works fine in any case.
        and not acu.FindInNavigatorItemTree(node, lambda i: i.navigatorItemId.guid in sourcesOfLinks)
    )
# The filtered list will contain only the 'father' not used elements itself.
unusedViewTreeItemsFiltered = []
# Loop through the 'unusedViewTreeItems' list
# to collect the 'father' elements from the 'unusedViewTreeItems' list.
for ii in unusedViewTreeItems:
    # Is it a child of an unused item? start value is False.
    isChildOfUnused = False
    # Loop through 'unusedViewTreeItems' list.
    for jj in unusedViewTreeItems:
        # If the actual element from the first loop is not the same as the actual element of the second loop
        # and their navigatorItemIds are equal, It means that it is a child element of a not used 'father element'.
        if ii != jj and acu.FindInNavigatorItemTree(jj, lambda node: node.navigatorItemId.guid == ii.navigatorItemId.guid):
            # Is it a child of an unused? Yes.
            isChildOfUnused = True
            # Leave it and go to the next element.
            break
    # If not False means it is True.
    if not isChildOfUnused:
        # It is a 'father' element so append to the filtered list.
        unusedViewTreeItemsFiltered.append(ii)
# Overwrite the old 'unusedViewTreeItems' list with the filtered one.
unusedViewTreeItems = unusedViewTreeItemsFiltered

# Rename the name of the items in the viewMapTree to folderName.
folderFromPreviousRun = acu.FindInNavigatorItemTree(viewMapTree.rootItem, lambda i: i.name == folderName)
# If 'folderFromPreviousRun' exist (not empty) and 'renameFolderFromPreviousRun' is True.
if folderFromPreviousRun and renameFolderFromPreviousRun:
    # Rename the first element of the 'folderFromPreviousRun' to folderNameForPreviousRun.
    acc.RenameNavigatorItem(folderFromPreviousRun[0].navigatorItemId, newName=folderNameForPreviousRun)

# The base case is None.
unusedViewsFolder = None

# If 'moveToFolder' is True.
if moveToFolder:
    # If 'renameFolderFromPreviousRun' is False and 'folderFromPreviousRun' is not empty.
    if not renameFolderFromPreviousRun and folderFromPreviousRun:
        # The unused views folder to be the folder with folderName.
        unusedViewsFolder = folderFromPreviousRun[0].navigatorItemId
    else:
        # Create unused views folder with the name of folderName.
        unusedViewsFolder = acc.CreateViewMapFolder(act.FolderParameters(folderName))

# Loop through the sorted 'unusedViewTreeItems'.
for item in sorted(unusedViewTreeItems, key=lambda i: i.prefix + i.name):
    # Try - except block not necessary.
    try:
        # If 'moveToFolder' and 'unusedViewsFolder' are True and exist.
        if moveToFolder and unusedViewsFolder:
            # Move the unused navigator item to the 'unusedViewsFolder'.
            acc.MoveNavigatorItem(item.navigatorItemId, unusedViewsFolder)
        # Print out onto the concole the item prefix, item.name and the item itself.
        print(f"{item.prefix} {item.name}\n\t{item}")
    except:
        continue
