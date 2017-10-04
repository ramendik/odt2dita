# -*- coding: utf8 -*-
# (C) IBM, Mikhail Ramendik 2010,2017
# Released as open source with IBM approval 

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# This is a BETA script to convert Open Document Text format documents
# into DITA
# This limited release version targets OASIS DITA 

# Its aim is to minimize the need for manual actions pre- or post conversion
# With any questions or suggestions regarding this script please contact:
# mr@ramendik.ru

# NOTES ON COMMENT TERMINOLOGY
# "cell/item" is a table cell, a list item, or anything similar - where a paragraph content is
#  used but not the <p> tag itself, unless there is more than one paragraph

# VERSION: 0.42 OASIS
# October 4, 2017

# TO DO LIST
# Support continuing an ordered list after an embedded list (apparently the text:continue-list attribute in text:list)
# Remove bold/italic formatting if it applies to an entire note

# CHANGE HISTORY
# 0.001 2010/04/14 Initial attempt
# 0.002 2010/04/16 first version that compiles and works successfully as coded:
#  - styles analyzed; supported properties: "bold", "italic", "monospace","note"
#  - text analyzed, text content output with style properties
#  - output in single file with provisional formatting
# Feasibility seems confirmed
# To do until first non-private test: graphics, tables, lists, headings (and multiple DITA with them)
# To do more until first milestone: cross-references. This includes auto ID displacement and
#  initial postprocessing (for references across files)
# 0.003 2010/04/16 now writing valid DITA!
# 0.004 2010/04/17 processing tables, images, text boxes (content only)
#                  also first postprocessing routine added - fixing nested paragraphs
#                  also can write out multiple DOMs (but not actually creating them yet)
#                  to do until first PDF: (at least)
#                        lists, headings, text:s, removal of extra whitespace and blank paragraphs
# 0.005 2010/04/22 now removes extra whitespace, processes text:s
#                  improved bold/italic/mono processing
#                  added optional uicontrol support from "antiqua" font
#                  added links (HTTP/WWW only)
#                  TODO: lists, headings, postprocessing including table stuff and extra blanks.
#                           and also caption handling - catch it in style mechanism and otherprops,
#                           then use in postprocessing
# 0.006 2010/04/22 now processes lists.
#                   IMPORTANT: in postprocessing we will need to join adjacent <ul> lists
# 0.007 2010/04/28 "header" and "caption" styles now marked in otherprops, needed for table preprocessing
#                   parent styles now handled correctly
#                   lots of table preprocessing done but some of it went wrong
# 0.008 2010/04/30 Working nearly fully - ironing out glitches
# 0.009 2010/04/30 First demo working
# 0.010 2010/04/30 bugfix
# 0.011 2010/04/30 Links apparently working
#                    TODO: check links between topics; process index terms; enhance table caption; remove <b>  </b> etc
# 0.012 2010/05/05 enhanced table caption, DOM IDs now made from title, "<b>  </b>" cases removed
#                   also added final postproc routines to remove and replace tags - now unused
# 0.013 2010/05/07 user requested improvements:
#    All file names and IDs are now lower case
#    Empty shortdesc paragraphs are added after every topic title
#    Topic type can be set in title text. Concept is the default. [r] means reference, [t] means task.
#      ATTENTION: task will have a dummy step added.
# 0.014 2010/05/13 fixed bug - duplicate style name caused end of style processing
# 0.015 2010/05/14 design change - moved breakup into topics to postprocessing
#                    before that, titles are saved as "temp:topic"
#                   reason: for Symphony support, we need to be able to process text:h within lists as new topic
#                   this means we also have to handle the cross-ref dictionary in postprocessing
#                   also removing (R) symbol (TM symbol not found in target docs as yet)
# 0.016 2010/05/26 processing a zip odt directly
# 0.017 2010/05/26 GUI started
# 0.018 2010/06/18 First working GUI; some bug fixes, make notes from paragraphs starting with "Note:" and similar
# 0.019 2010/07/23 GUI ready for user handover (though would still want some improvement)
#        Transition from alpha to beta should be somewhere soon
# 0.020 2010/07/25 file and dir selection dialogs made prettier
#       ?? make output a file selection so one can create a dir?
# 0.021 2010/08/12 warning not crash on broken link; support for footnotes
# 0.022 2010/10/17 fixed a rare crash when a style name is used but not defined
#                   added experimental Frame Mode
# 0.023 2010/10/22 fixed a bug in Frame Mode - global dictionary was not reset between document conversions 
# 0.024 2010/10/27 removing [c] tag. Also, changing image file extensions - png to jpg, all other non-jpg to gif
# 0.025 2010/11/04 image ext changing now fully works, added format="dita" in map creation
# 0.026 2010/11/30 conc_ etc changed to c_ etc ; tag replacement added
# 0.027 2011/02/17 processing headers that have outline-level in style but no text:h (COMMENTED OUT for now);
#                   bug fix - correct processing of the "id" attribute of a node to be destroyed
#                   Thanks to Cristina Bonanni for the testing that has led to
#                   these improvements
# 0.28 2011/03/29 Fixed a typo bug that prevented list styles processing
#                   Thanks, again, to Cristina Bonanni for the testing
#                   Also removed a zero from the version as the script is now in beta, not alpha
# 0.29 2011/04/07 Added support for header columns listing
#                 Added removing blank and whitespace-only notes
#                 Added a fix for nested lists
#                 Added a To Do List in the beginning
#                 Everything thanks to Cristina Bonanni!
# 0.30 2011/04/15 With InTable>0, text:h now treated as text:p
#                   in PostprocessTables, handle the case when <tbody> is not found without a crash
#                   Again thanks to Cristina Bonanni
# 0.31 2011/07/07 Fixed a bug that led to empty output with some Symphony docs
# 0.32 2012/06/28 Fixed a bug that led to exceptions in rare cases
# 0.33 2012/07/20 Fixed rare exception with a tricky bookmarks case
# 0.34 2013/02/02 Fixed rare exception when a "dc:title" node exists but does not have text in it
# 0.35 2013/02/28 Removed PIL dependency by procesing PNG files as PNG
# 0.36 2013/05/15 Fixed multiple _ in naming, added option to not prefix
# 0.37 2013/06/02 Added optional task step processing, fixed tag renaming
# 0.38 2014/01/06 Added MathML processing (if there is MathML in source, output now requires IDWB 4.5.0)
# 0.39 2014/06/26 Worked around obscure Python set problem that sometimes led to extra italics; removed non-ascii chars in filenames/ids; join adjacent i/u/codeph
# 0.40 2014/11/07 Worked around a case when a style name is not defined but used (looks like *some* cases of this problem were covered earlier too)
#                    Also enabled step processing in tasks by default
# 0.41 2014/12/09 Aggressive formula detection mode
# 0.41 OASIS 1 2016/10/10 OASIS branch
# 0.42 OASIS 2017/10/04 open source, command line interface, removed skipping boilerplate as it is confusing on small documents, added "notitle"

import xml.dom
import xml.dom.minidom
import sys
import os
import os.path
import zipfile
from Tkinter import *
from Tix import *
import tkMessageBox
import tkFont
import tkFileDialog
import codecs
import copy
import argparse




# INPUT DATA
# This should be replaced with a nice graphical input window
# And an unzip routine from the original ODT file should be added
# For now, manually unzip and set InputDirectory to the full path to unzipped ODT

InputFile="NOT SELECTED"

OutputDirectory="NOT SELECTED"

# Frame Mode - use "headingN" style names as heading indicators

# Skip the first amount of output sections
# Useful for skipping boilerplate
SkipOutputSections=0

# Tricky processing options
Process_AntiquaAsBold=True
Process_DeleteBadLinks=False

# Root part for automatic naming; used for IDs, and then DOM IDs are used for file names
NameRoot=""

# Level of debug output
DebugLevel=3

# Options added in 0.36
#DoNotPrefix=False
#UsePrettyXml=False

# Added in 0.41
AggressiveFormula=False

# GLOBAL DATA

# Document title - to be worked out from meta.xml
DocumentTitle=""

# Bookmarks dictionary
# Key is the bookmark name, as in document
# Member is a tuple (DocID,reference_in_document)
BookmarksDictionary={}

# Duplicate bookmarks
# While the bookmark is normally just saved as the ID of the parent paragraph or title, there can be two bookmarks for one
# This dictionary is used in this case
BookmarksDuplicate={}

# Styles dictionary
# Key is the style name, and member is a set of keywords with the style properties
# Style properties defined as of this version: "bold", "italic", "monospace","note"
# "notbold","notitalic","notmonospace","uicontrol","caption","header"
StylesDictionary={"":set([])}


# Outline levels for styles
#StylesOutlineLevels={"":0}

# This counter is increased when starting to process a table, decreased when ending the table.
# !!!!!!!!!! this is not a proper solution !!!!!!!!!!!!!
InTable=0

# Frame Mode - use "headingN" style names as headings
FrameMode=False

# Task postprocessing - attempt to create steps
TaskPost=False

# Replacement tags for <b> and <i>, empty if no replacement is required
ReplaceBoldWith=""
ReplaceItalicWith=""

# Frame Mode styles dictionary
# Used only in Frame Mode
# Contains only styles with "headingN" in name or if the parent style is in the list
# The value is an integer - the heading level

FrameModeStylesDictionary={}

# Style tags
# This is a dictionary of tags by which style properties are represented
# At present this is static
# NOTE: for "off" styles this is the tag thagt is to be turned off
# NOTE: for "otherprops" styles this is the string added to the otherprops attribute
StyleTags={"bold":"b","italic":"i","monospace":"codeph","note":"note",
           "notbold":"b","notitalic":"i","notmonospace":"codeph","uicontrol":"uicontrol",
           "caption":"caption","header":"header"}

# Which style properties are character styles, that is, can be "switched on"
# inside a paragraph
StylesCharacter=set(["bold","italic","monospace","uicontrol"])

# Which style properties are character "off" styles, that is, can "switch off"
# character styles inside a paragraph
StylesCharacterOff=set(["notbold","notitalic","notmonospace"])


# Which style properties are paragraph styles. These must be processed before character styles
StylesParagraph=set(["note"])

# Which style properties are "otherprops" styles. They are paragraph level styles and
# reflected by adding a string to the "otherprops" attribute of the <p> or other parent tag,
# for use in postprocessing
StylesOtherprops=set(["caption","header","monospace"])

# List styles are handled separately because each has a number of levels
# We take a "dictionary of dictionaries" approach
# Each member in the ListStylesDictionary is a dictionary of level to string
# String can be "number" or "bullet"
ListStylesDictionary={"":{}}

# Current list level and stack of list style names
# FIXME: use of globals is probably not a good idea
CurrentListLevel=0
ListStyleNamesStack=[""]

# Output DOMs
# This is a list of tuples (DOM, level)
# DOM is the output DOM
# level is the outline level
# note we do not save DOM IDs separately, to avoid complicated data structures
OutputDOMs=[]

# Automatically created IDs
# We often need to replace the ID of a topic - other tags too perhaps in later versions
# that was first created as automatic. In such cases we need to ensure that the ID is indeed
# automatic. If the ID to be replaced is already from the document we have a non-standard situation
AutoCreatedIDs=[]

# All IDs, auto-created and otherwise
# This is here to catch non-unique IDs, which should not happen anyhow
# A non-unique ID may lead to a non-unique file name and other glitches, so is a non-standard situation
AllIDs=[]

# ID counter. For use in autocreating IDs; increases each time an ID is autocreated
# Initiated with 0 because the very first output DOM is somewhat likely to be empty
#  (and therefore not output to a DITA file)
AutoID_Counter=0

# 0.38 additional files to unzip
ToUnzip=set([])

# Tags to ignore
# Tags that, when found in text or other places, are to be ignored
#  (and not logged as unprocessed)
# FIXME: split this by place where to ignore them?
TagsToIgnore=["text:table-of-content",  # Ignore the TOC
              "text:sequence-decls",
              "office:forms",
              "text:bookmark-end"]

# Current output DOM and its <body>
# FIXME: keeping the current output DOM global is not a good solution - only a simple one
InitialOutputDOM=None
InitialOutputBody=None
CurrentOutputTitle=None

# Current output DOM's <keywords>
# FIXME: do not use a global?
# set to None until the thing is actually created
CurrentOutputKeywords=None

# === SERVICE FUNCTIONS ====================

# Debug output, with its global
DebugText=""

def debug(level,message):
    global DebugText
    if level<=DebugLevel:
        DebugText=DebugText+message+"\n"

# Report an unprocessed tag in a parent tag, unless it's to be ignored
def unprocessed(parenttag,tag):
    if not (tag in TagsToIgnore):
        debug(3,"Unprocessed tag in "+parenttag+" : "+tag)

# copy attributes from input node to output node - assume both are elements
def CopyAttributes(InputNode,OutputNode):
    for i in range(0,InputNode.attributes.length):
        attr=InputNode.attributes.item(i)
        OutputNode.setAttribute(attr.nodeName,attr.value)




# Carefully destroy a node and all its child nodes - remove from tree and unlink
# Save the ID (or any child node ID) to the parent node
def DestroyNode(node):
    while node.firstChild<>None:
        #print "destroying a child node"
        DestroyNode(node.firstChild)

    SaveId(node)
    node.parentNode.removeChild(node)
    node.unlink()

# Find the first text node under the node
# If the node is text or None return itself; if the none is non-text and non-element/doc return None
# If the node has no text subnodes return None
def FindFirstText(node):
    if node==None:
        return None
    elif node.nodeType==xml.dom.Node.TEXT_NODE:
        return node
    elif node.nodeType in [xml.dom.Node.ELEMENT_NODE,xml.dom.Node.DOCUMENT_NODE]:
        result=None
        for childnode in node.childNodes:
            result=FindFirstText(childnode)
            if result<>None: break
        return result
    else:
        return None

# Find the LAST text node under the node
# If the node is text or None return itself; if the none is non-text and non-element/doc return None
# If the node has no text subnodes return None
def FindLastText(node):
    if node==None:
        return None
    elif node.nodeType==xml.dom.Node.TEXT_NODE:
        return node
    elif node.nodeType in [xml.dom.Node.ELEMENT_NODE,xml.dom.Node.DOCUMENT_NODE]:
        result=None
        for childnode in node.childNodes:
            newresult=FindFirstText(childnode)
            if newresult<>None:
                result=newresult
        return result
    else:
        return None


# Move all child nodes from InputNode to OutputNode, to follow its existing children
# Merge text nodes (this is done by built-in normalize method)
# If TextSeparator is present, place the text between the original children and the moved ones
# NOTE: does not delete InputNode even though strips it of all children
def MoveChildNodes(InputNode,OutputNode,TextSeparator=''):

    # add TextSeparator    
    if TextSeparator<>"":
        newnode=InitialOutputDOM.createTextNode(TextSeparator)
        OutputNode.appendChild(newnode)

    # move nodes
    while InputNode.hasChildNodes():
        OutputNode.appendChild(InputNode.firstChild)

    # Normalize text nodes

    OutputNode.normalize()


# Get the element child of a DOM
def GetElementChild(DOM):
    for childnode in DOM.childNodes:
        if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
            return childnode
    debug(1,"GetElementChild failed to get element child")
    return None
            
        


# Get the ID from an output DOM
def GetID(DOM):
    topicnode=GetElementChild(DOM)
    if topicnode.nodeType<>xml.dom.Node.ELEMENT_NODE:
        debug(1,"GetID: non-element last child")
        return("")

    #FIXME: here we check for the fact that this is the root element of a valid DITA doc
    #but really valid DITA documents can have several root element not just <topic>
    #for now we only create <topic> but we do need to check properly here
    if not (topicnode.tagName in ["topic","concept","task","reference"]):
        debug(1,"GetID: non-DITA first child (not topic type")
        return("")

    return(topicnode.getAttribute("id"))

# Set the ID for an output DOM
def SetID(DOM,idstring):
    topicnode=GetElementChild(DOM)

    #FIXME: here we check for the fact that this is the root element of a valid DITA doc
    #but really valid DITA documents can have several root element not just <topic>
    #for now we only create <topic> but we do need to check properly here
    if not (topicnode.tagName in ["topic","concept","task","reference"]):
        debug(1,"SetID: non-DITA first child (not topic type)")
        return("")

    topicnode.setAttribute("id",idstring)

# Get the title node for an output DOM
def GetTitleNode(DOM):
    topicnode=GetElementChild(DOM)
    
    if not (topicnode.tagName in ["topic","concept","task","reference"]):
        debug(1,"GetTitleNode: non-DITA first child (not topic type")
        return("")

    for childnode in topicnode.childNodes:
        if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
            if childnode.tagName=="title":
                return childnode
    debug(1,"GetTitleNode: did not find title node")
    return None
            

# Create, and return, a new <temp:topic> in the INITIAL output DOM, append to node
# New for version 0.015
# The <temp:title> has the <title> node as well as the <prolog> set up for the keywords

def CreateNewTempTopic(node):
    global IsStartParagraph
    IsStartParagraph=True

    TTNode=InitialOutputDOM.createElement("temp:topic")
    node.appendChild(TTNode)

    TitleElement=InitialOutputDOM.createElement("title")
    TTNode.appendChild(TitleElement)
    global CurrentOutputTitle
    CurrentOutputTitle=TitleElement

    prolognode=InitialOutputDOM.createElement("prolog")
    TTNode.appendChild(prolognode)
    metadatanode=InitialOutputDOM.createElement("metadata")
    prolognode.appendChild(metadatanode)
    global CurrentOutputKeywords
    CurrentOutputKeywords=InitialOutputDOM.createElement("keywords")
    metadatanode.appendChild(CurrentOutputKeywords)

    return TTNode


    
            
# Create the INITIAL output DOM
# This DOM is used for main processing and most postprocessing
# It has only the body node, and is later "broken up" into actual output DOMs
# New for version 0.015

def CreateInitialOutputDOM():
    global InitialOutputDOM
    impl=xml.dom.getDOMImplementation("minidom")
    InitialOutputDOM=impl.createDocument(None,"conbody",None)
    global InitialOutputBody
    InitialOutputBody=GetElementChild(InitialOutputDOM)
    CreateNewTempTopic(InitialOutputBody)
    






# Remove all childless element children of a node
# This is to facilitate things like paragraph processing: there may be lots of bold etc tags
#  but if there is no content all such tags should be removed
# All child elements are processed recursively, but all non-element children remain
# ONLY elements from a certain list - "b","i","codeph","p","note" - are removed

def RemoveChildlessElementChildrenPara(node):
    # child nodes flagged for removal
    ChildNodesToRemove=[]

    for childnode in node.childNodes:
        
        # processing a child node now
        # only interested if it is an element
        if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
            # first remove any of its own element children by recursion 
            RemoveChildlessElementChildrenPara(childnode)

            # now if this child node is childless, flag it for removal
            # not removing immediately just to avoid trouble with the for statement
            if not childnode.hasChildNodes():
                if childnode.tagName in ["b","i","codeph","p","note"]:
                    ChildNodesToRemove.append(childnode)

    # remove flagged child nodes, and return
    for childnode in ChildNodesToRemove:
        SaveId(childnode)
        node.removeChild(childnode)
        childnode.unlink()
    return()
            
# The paragraph start mark for ignoring white space at start of paragraph
# FIXME: using a global for this is probably not a good way
IsStartParagraph=True

# Add some text as text child data to OutputNode
# Processes white space, and does not create another text child is the last child is text already

def AddTextAsChild(text,OutputNode):

    global IsStartParagraph
    
    # if text is empty, simply return as we have nothing to add
    if text=="":
        return()

    # determine if the last child of OutputNode is text, and if so does it end with white space 
    # also, if this is the paragraph start, beginning spaces are to be removed as if after a space
    wasSpace=IsStartParagraph 
    ExistingText=False
    if OutputNode.lastChild <> None:
        if OutputNode.lastChild.nodeType==xml.dom.Node.TEXT_NODE:
            ExistingText=True
            wasSpace=OutputNode.lastChild.data[len(OutputNode.lastChild.data)-1].isspace()

        

    # process the text data, shrinking all white space

    data=text
    newdata=""
    while data<>"":
        # process data[0]
        if data[0].isspace():
            # add a space if this is the first white space
            if not wasSpace:
                newdata=newdata+" "
                wasSpace=True
        else:
            wasSpace=False
            newdata=newdata+data[0]
        # now remove data[0]
        data=data[1:]

        
        
    # if resulting text is empty, simply return as we have nothing to add
    if newdata=="":
        return()

    # add resulting text
    if ExistingText:
        OutputNode.lastChild.data=OutputNode.lastChild.data+newdata
    else:
        newnode=InitialOutputDOM.createTextNode(newdata)
        OutputNode.appendChild(newnode)

    IsStartParagraph=False
    return()


# Add number spaces as text child data to OutputNode
# Does not create another text child is the last child is text already
# And supports AddTextAsChild globals

def AddSpacesAsChild(number,OutputNode):

    global IsStartParagraph
    
    # determine if the last child of OutputNode is text
    
    ExistingText=False
    if OutputNode.lastChild <> None:
        if OutputNode.lastChild.nodeType==xml.dom.Node.TEXT_NODE:
            ExistingText=True

    newdata=""
    for i in range(0,number):
        newdata=newdata+" "

    # add resulting text
    if ExistingText:
        OutputNode.lastChild.data=OutputNode.lastChild.data+newdata
    else:
        newnode=InitialOutputDOM.createTextNode(newdata)
        OutputNode.appendChild(newnode)

    IsStartParagraph=False
    return()



# == PROCESSING ==========================================================

# process a styles list node
# The styles list is either <office:automatic-styles> in content.xml
#  or <office:styles> and <office:automatic-styles> in styles.xml
# All element children of the styles list node are to be fed to this procedure
def ProcessStylesNode(InputNode):
    global StylesDictionary
    global ListStylesDictionary
#    global StylesOutlineLevels

#    print "PSN started"
    
    if InputNode.nodeType<>xml.dom.Node.ELEMENT_NODE:
        debug(1,"ProcessStyleNode: non-element node received")
        return()

    for node in InputNode.childNodes:
      if node.nodeType==xml.dom.Node.ELEMENT_NODE:
        if node.tagName=="style:style": #process para/char/etc styles
            StyleName=node.getAttribute("style:name")
            
            if StyleName in StylesDictionary:
                # do not process duplicate style name
                debug(1,"ProcessStyleNode: duplicate style "+StyleName)
                continue
              

            # note: in the future we might distinguish by style:family attribute here
            # not yet needed at this stage so proceed straight to working out properties

            # initialize style properties 
            StyleProperties=set([])

            # first checks for note, header and caption

            if StyleName.lower().find("note")>-1:
                StyleProperties.add("note")
            if StyleName.lower().find("head")>-1:
                StyleProperties.add("header")
            if StyleName.lower().find("caption")>-1:
                StyleProperties.add("caption")

            # get the parent style name, and add parent style attributes and outline level if present
            parentstyle=node.getAttribute("style:parent-style-name")
            try:
                StyleProperties = StyleProperties | StylesDictionary[parentstyle]
            except KeyError: pass  
#            try:
#                StylesOutlineLevels[StyleName]=StylesOutlineLevels[parentstyle]
#            except KeyError: StylesOutlineLevels[StyleName]=0

            # get the outline level for the style
#            try:
#                if node.hasAttribute("style:default-outline-level"):
#                    StylesOutlineLevels[StyleName]=int(node.getAttribute("style:default-outline-level"))
#                    print "Style "+StyleName+" outline level "+str(StylesOutlineLevels[StyleName])
#            except ValueError:pass

            
            # do second checks for note/header/caption
            if parentstyle.lower().find("note")>-1:
                StyleProperties.add("note")
            if parentstyle.lower().find("head")>-1:
                StyleProperties.add("header")
            if parentstyle.lower().find("caption")>-1:
                StyleProperties.add("caption")

            # Process heading styles for Frame Mode
            if FrameMode:
                if parentstyle in FrameModeStylesDictionary:
                    FrameModeStylesDictionary[StyleName]=FrameModeStylesDictionary[parentstyle]
                    debug(3,"found Frame heading child style: "+StyleName)
                elif StyleName.lower().find("heading")>-1:
                    LevelStr=StyleName[StyleName.lower().find("heading")+7:]
                    try:
                        nn=int(LevelStr)
                        FrameModeStylesDictionary[StyleName]=nn
                        debug(3,"found Frame heading level "+LevelStr+" style name "+StyleName)
                    except ValueError: pass # no number means not a Frame heading
                    




            # other style properties are in child elements whose names include "properties"
            # so we cycle through all child elements finding such nodes
            # and check their attributes for the properties that we detect

            for childnode in node.childNodes:
                if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                    childTagName=childnode.tagName
                    if childTagName.find("properties")>-1:
                        # this is a properties element - look for properties

                        # look for bold and notbold
                        if childnode.getAttribute("fo:font-weight").find("bold")>-1:
                            StyleProperties.add("bold")
                        if childnode.getAttribute("fo:font-weight").find("normal")>-1:
                            StyleProperties.add("notbold")
                            StyleProperties.discard("bold") # if it was in parent

                        
                        # look for italic and notitalic
                        if childnode.getAttribute("fo:font-style").find("italic")>-1:
                            StyleProperties.add("italic")

                        if childnode.getAttribute("fo:font-style").find("normal")>-1:
                            StyleProperties.add("notitalic")
                            StyleProperties.discard("italic") # if it was in parent
                            
                        # look for monospace and notmonospace
                        # FIXME: is there a better way than checking for "Cour*" fonts?
                        fontname=childnode.getAttribute("style:font-name").lower()
                        if fontname<>"":
                            if fontname.find("cour")>-1:
                                StyleProperties.add("monospace")
                            else:
                                StyleProperties.add("notmonospace")
                                StyleProperties.discard("monospace") # if it was in parent
                            global Process_AntiquaAsBold
                            if Process_AntiquaAsBold and fontname.find("antiqua")>-1:
                                StyleProperties.add("bold")
                            
            # Save the style properties in the dictionary
#            print StyleName
            StylesDictionary[StyleName]=StyleProperties
        elif node.tagName in ["text:liststyle","text:list-style"]: #process list styles
            #print "Found list styles"
            StyleName=node.getAttribute("style:name")
            StyleLevels={}

            # list levels are in child elements - text:list-level-style-bullet
            # and text:list-level-style-number
            
            for childnode in node.childNodes:
                if childnode.tagName=="text:list-level-style-bullet":
                    StyleLevels[int(childnode.getAttribute("text:level"))]="bullet"
                if childnode.tagName=="text:list-level-style-number":
                    StyleLevels[int(childnode.getAttribute("text:level"))]="number"

            ListStylesDictionary[StyleName]=StyleLevels


                            
# process a paragraph - presumably the contents of a <text:p> node
# such a node is passed but we do NOT check it in any way, so it can be a different tag
#  (this allows using the function for <text:h> when heading processing is not enabled)
# returns True if any result was rendered into the output, False otherwise
# Note that a <p> is NOT created; if a <p> is neded the OutputNode should be it
#  (this is because it can be the first paragraph of a cell/item)
# ParagraphStyleProperties is a style properties set
# If IsTitle=True, this is a <text:h> node going into a title, ignore paragraph style
#  and restrictions apply

def ProcessParagraph(InputNode,OutputNode,ParagraphStyleProperties=set([]),IsTitle=False):

    # LOCAL FUNCTIONS OPERATING AN OUTPUT NODE STACK
    # output node stack: get current output node
    def CurrentOutputNode():
        return(OutputNodes[len(OutputNodes)-1])
    
    # output node stack: add a node with the given tag name as output node, add to stack
    # FIXME: add attribute support?
    def StartOutputNode(tag):
        newnode=InitialOutputDOM.createElement(tag)
        CurrentOutputNode().appendChild(newnode)
        OutputNodes.append(newnode)

    # output node stack: end the current output node, use the previous one
    def EndOutputNode():
        OutputNodes.pop()

    # output node stack: peel the output stack until we reach peelnode or (error!)
    # the original OutputNode
    def PeelOutputStack(peelnode):
        while not CurrentOutputNode() in [peelnode,OutputNode]:
            EndOutputNode()
        if CurrentOutputNode()<>peelnode:
            debug(3,"PeelOutputStack called with a peelnode not in the stack")

    # Set (only to true - this is a one way trigger) and get peeled flag
    # FIXME: this is a hack, but namespace issues happen otherwise
    def SetPeeled():
        Peeled.append(1)
    def GetPeeled():
        if Peeled==[]:
            return(False)
        else:
            return(True)

    # LOCAL FUNCTION FOR PROCESSING TEXT
    # Required because <text:span> can be nested
    # Process all children of LocalInputNode into the current output node as per stack
    # The style properties (character only are processed) are considered already set up
    #  as by LocalStyleProperties
    def LocalProcessText(LocalInputNode,LocalStyleProperties,DebugPrints=False):
        if DebugPrints:
            print "LocalProcessText called with DebugPrints"
            print LocalStyleProperties
        for childnode in LocalInputNode.childNodes:
            # process plain text by copying it
            if childnode.nodeType==xml.dom.Node.TEXT_NODE:
                AddTextAsChild(childnode.data,CurrentOutputNode())
                
            # process child elements
            # FIXME: this currently is NOT verified to process every possible child as per spec!
            #  (unprocessed child nodes will be logged at debug level 3)
            if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                if DebugPrints:
                    print "processing node "+childnode.tagName+" with text:style-name: "+childnode.getAttribute("text:style-name")
                    print LocalStyleProperties
                
                if childnode.tagName in ["text:description","text:sequence"]:
                    # process contents as text
                    LocalProcessText(childnode,LocalStyleProperties)
                elif childnode.tagName=="text:span":
                    # get new style
                    try:
                        newstyleproperties=StylesDictionary[childnode.getAttribute("text:style-name")]
                    except KeyError:
                        # a rare ODT issue - style not defined
                        #!!TEMP
                        #print "not found"
                        newstyleproperties=LocalStyleProperties
                    if DebugPrints:
                        print LocalStyleProperties

                    # form the real style properties - start with existing ones 
                    styleproperties=copy.deepcopy(LocalStyleProperties)
                    # apply the "off" properties of the new style
                    if "notbold" in newstyleproperties:
                        styleproperties.discard("bold")
                    if "notitalic" in newstyleproperties:
                        styleproperties.discard("italic")
                    if "notmonospace" in newstyleproperties:
                        styleproperties.discard("monospace")
                    # now add any new character properties
                    styleproperties=styleproperties | (newstyleproperties & StylesCharacter)

                    if DebugPrints:
                        print LocalStyleProperties


                    # !!TEMP
                    if DebugPrints and childnode.getAttribute("text:style-name")=="T127":
                        print "T127"
                        print styleproperties
                        print "local is"
                        print LocalStyleProperties

                        
                    # now we have the new style in styleproperties
                    # first, check if any character style is currently on, but off in new style
                    # and depending on this, ensure we set up the character style correctly

                    # IMPORTANT: we are using a halfway house approach here
                    # If the stack was ever peeled back to reset all character props
                    #  then we just have to do it every single time within this paragraph,
                    #  but this only happens if some prop was in a top level tag but not
                    #  a bottom level one - a rare situation

                    if ((LocalStyleProperties-styleproperties) & StylesCharacter)<>set([]) or GetPeeled():
                        # We need to peel stack and re-setup all character style properties

                        #!!TEMP
                        #print "peeling"


                        SetPeeled()
                        PeelOutputStack(CharacterRoot)
                        for style in (styleproperties & StylesCharacter):
                                StartOutputNode(StyleTags[style])
                    else:
                        # no character style property removal
                        # just add any missing ones
                        # and save the local root to return to after processing the text

                        LocalRoot=CurrentOutputNode()
                        for style in ((styleproperties-LocalStyleProperties) & StylesCharacter):
                            StartOutputNode(StyleTags[style])

                    # now, use recursion to process the contents of the span
                    dp=False
                    #if childnode.getAttribute("text:style-name")=="_7e_Special_20_term":
                    #    dp=True
                    LocalProcessText(childnode,styleproperties,dp)

                    # get the old character style properties back before continuing
                    if GetPeeled():
                        # peel to character root and reset properties fully
                        PeelOutputStack(CharacterRoot)
                        for style in (LocalStyleProperties & StylesCharacter):
                            StartOutputNode(StyleTags[style])
                    else:
                        # just get back to the local root, which still has the old properties
                        PeelOutputStack(LocalRoot)
                elif childnode.tagName=="draw:image" and not IsTitle: # process an image by creating an image tag
                    imagenode=InitialOutputDOM.createElement("image")
                    image_href=childnode.getAttribute("xlink:href")
                    (image_root,image_ext)=os.path.splitext(image_href)
                    if not (image_ext.lower() in [".jpg",".png"]):
                        image_ext=".png"
                    imagenode.setAttribute("href",image_root+image_ext)
                    imagenode.setAttribute("placement","break")
                    CurrentOutputNode().appendChild(imagenode)
                    debug(3,"Image original "+image_href+" set as "+image_root+image_ext)
                elif childnode.tagName.find("draw:")>-1 and not IsTitle: # draws other than image - process content as text node

                    # 0.38 check for formula
                    
                    foundFormula=False
                    for grandchildnode in childnode.childNodes:
                        if grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                            if grandchildnode.tagName=="svg:desc":
                                descnode=grandchildnode.firstChild
                                if descnode.nodeType==xml.dom.Node.TEXT_NODE:
                                    if descnode.data.strip().lower()=="formula":
                                        foundFormula=True
           
                    
                    ProcessTextNode(childnode,False,OutputNode,foundFormula)
                    # IMPORTANT: this may cause nested <p> tags!
                    # fixed in postprocessing, to avoid further code clutter
                elif childnode.tagName=="text:note" and not IsTitle:
                    # process a footnote, or an endnote that becomes a footnote anyway
                    # make an <fn> tag and process the footnote body into it
                    fnnode=InitialOutputDOM.createElement("fn")
                    CurrentOutputNode().appendChild(fnnode)
                    for grandchildnode in childnode.childNodes:
                        if grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                            if grandchildnode.tagName=="text:note-body":
                                ProcessTextNode(grandchildnode,False,fnnode)
                elif childnode.tagName=="text:s": # add spaces
                    try:
                        number=int(childnode.getAttribute("text:c"))
                    except ValueError:
                        number=1
                    AddSpacesAsChild(number,CurrentOutputNode())
                elif childnode.tagName=="text:tab": # FIXME: add 6 spaces - this may not be correct
                    AddSpacesAsChild(6,CurrentOutputNode())
                elif childnode.tagName=="text:a":
                    # FIXME: only URI links for now
                    # but nonprocessed links get text processed anyhow
                    # if IsTitle, only process text anyway

                    LocalRoot=CurrentOutputNode()
                    linktarget=childnode.getAttribute("xlink:href")
                    if (linktarget.lower().find("://")>-1 or linktarget.lower().find("mailto")>-1) \
                       and not IsTitle:
                        newnode=InitialOutputDOM.createElement("xref")
                        CurrentOutputNode().appendChild(newnode)
                        OutputNodes.append(newnode)
                        newnode.setAttribute("scope","external")
                        newnode.setAttribute("format","html")
                        newnode.setAttribute("href",linktarget)
                    LocalProcessText(childnode,LocalStyleProperties)
                    PeelOutputStack(LocalRoot)                    
                elif childnode.tagName=="text:line-break" and not IsTitle:
                    # Processing a line break here would entail a new paragraph - too complicated
                    # So we insert a <temp:linebreak> element and rely on postprocessing
                    newnode=InitialOutputDOM.createElement("temp:linebreak")
                    CurrentOutputNode().appendChild(newnode)
                elif childnode.tagName in ["text:bookmark","text:bookmark-start"]:
                    BookmarkID=childnode.getAttribute("text:name")
                    #print "processing bookmark"
                    
                    # Bookmark - first save it dictionary

                    # saving REMOVED for 0.015
                    #DocID=GetID(InitialOutputDOM)
                    #Ref="#"+DocID
                    #if not IsTitle:  # if this is a title, ref will be to whole doc, otehrwise to ID in doc
                    #    Ref=Ref+"/"+BookmarkID
                    #BookmarksDictionary[BookmarkID]=(DocID,Ref)

                    # set the ID of the parent paragraph, or if it already exists, save duplicate
                    if OutputNode.getAttribute("id")=="":
                        OutputNode.setAttribute("id",BookmarkID)
                    else:
                        BookmarksDuplicate[BookmarkID]=OutputNode.getAttribute("id")
                        #print "duplicate added at para proc: "+BookmarksDuplicate[BookmarkID]
                        
                elif childnode.tagName=="text:bookmark-ref":
                    # Create a reference to the bookmark, with scope=local
                    # NOTE: the actualy reference will be changed in postprocessing
                    # FIXME? Currently we do NOT process text inside the ref, as DITA autogenerated link text
                    #  is usually better
                    newnode=InitialOutputDOM.createElement("xref")
                    CurrentOutputNode().appendChild(newnode)
                    newnode.setAttribute("scope","local")
                    newnode.setAttribute("href",childnode.getAttribute("text:ref-name"))
                elif childnode.tagName=="text:alphabetical-index-mark":
                    # need to create an indexterm - in prolog keywords if IsTitle, on top of paragraph otherwise
                    indextext=childnode.getAttribute("text:string-value")
                    indexkey1=childnode.getAttribute("text:key1")
                    # FIXME key2 not supported at present
                    
                    global CurrentOutputKeywords
                    if indextext<>"":
                        if IsTitle:
                            nodeforindexterm=CurrentOutputKeywords
                        else:
                            nodeforindexterm=OutputNode
                        indextermnode=InitialOutputDOM.createElement("indexterm")
                        nodeforindexterm.insertBefore(indextermnode,nodeforindexterm.firstChild)
                        if indexkey1<>"":
                            textnode=InitialOutputDOM.createTextNode(indexkey1)
                            indextermnode.appendChild(textnode)
                            nodeforindexterm=indextermnode
                            indextermnode=InitialOutputDOM.createElement("indexterm")
                            nodeforindexterm.appendChild(indextermnode)
                        textnode=InitialOutputDOM.createTextNode(indextext)
                        indextermnode.appendChild(textnode)
                        
                                
                else:
                    unprocessed(LocalInputNode.tagName,childnode.tagName)

    # MAIN FUNCTION BODY STARTS HERE
    # set flag for paragraph start
    global IsStartParagraph
    IsStartParagraph=True

    # Set up the output node stack  
    OutputNodes=[OutputNode]

    # Set up the peeled flag
    Peeled=[]

    # if IsTitle, ignore any passed style properties
    if IsTitle:
        ParagraphStyleProperties=set([])
    
    # process otherprops styles
    # FIXME?: we put the otherprops in the parent tag without checking what tag it is
    # while it probably can't break DITA or postprocessing, if anything does go wrong,
    # check this as a possible reason
    
    for style in (ParagraphStyleProperties & StylesOtherprops):
        OutputNode.setAttribute("otherprops",OutputNode.getAttribute("otherprops")+StyleTags[style])

    
    # process paragraph styles
    for style in (ParagraphStyleProperties & StylesParagraph):
        StartOutputNode(StyleTags[style])
            
    # node for peeling the stack to if we need to reset all character style properties
    CharacterRoot=CurrentOutputNode()
    
    # process character styles
    for style in (ParagraphStyleProperties & StylesCharacter):
        StartOutputNode(StyleTags[style])

    # process all text
    LocalProcessText(InputNode, ParagraphStyleProperties)

    # Remove all resulting empty elements
    RemoveChildlessElementChildrenPara(OutputNode)
    
# Process a table
# InputNode is assumed to be a <table:table> element
# Important: no table title is created, and no table head is used - all rows are in body
# We assume any table head will be detected via otherprops, in postprocessing
# This is to be fixed in postprocessing
def ProcessTableNode(InputNode,OutputNode):

    # Internal routine to process a row node
    # Returns the table body node. This allows creation of the body within the routine
    def ProcessRow(InputNode,IsHeaderRow=False):
        # Table data has started. Create table core elements if they weer not created before
        if TableStarted==[]:
            TableStarted.append(1)
            tablenode=InitialOutputDOM.createElement("table")
            OutputNode.appendChild(tablenode)
            
            tgroupnode=InitialOutputDOM.createElement("tgroup")
            tablenode.appendChild(tgroupnode)
            tgroupnode.setAttribute("cols",str(ColumnCount))

            for i in range (1,ColumnCount+1):
                colspecnode=InitialOutputDOM.createElement("colspec")
                tgroupnode.appendChild(colspecnode)
                colspecnode.setAttribute("colname","col"+str(i))

            # Create the TBody element. All rows are to be created in it.
            body=InitialOutputDOM.createElement("tbody")
            tgroupnode.appendChild(body)
        else:
            body=TBodyNode
        
        RowNode=InitialOutputDOM.createElement("row")
        body.appendChild(RowNode)

        # process all cells
        CurrentColumn=1
        for grandchildnode in InputNode.childNodes:
            if grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                if grandchildnode.tagName=="table:table-cell":
                    EntryNode=InitialOutputDOM.createElement("entry")
                    RowNode.appendChild(EntryNode)
                    if IsHeaderRow:
                        EntryNode.setAttribute("otherprops","header")
                    spanned_string=grandchildnode.getAttribute("table:number-columns-spanned")
                    if spanned_string=="":
                        EntryNode.setAttribute("colname","col"+str(CurrentColumn))
                        CurrentColumn=CurrentColumn+1
                    else:
                        spanned=int(spanned_string)
                        EntryNode.setAttribute("namest","col"+str(CurrentColumn))
                        EntryNode.setAttribute("nameend","col"+str(CurrentColumn+spanned-1))
                        EntryNode.setAttribute("align","center")
                        # FIXME? Alignment should be set as in table style?
                        # but present approach is actually better for DITA uniformity
                        CurrentColumn=CurrentColumn+spanned                          

                    ProcessTextNode(grandchildnode,False,EntryNode)
        return body

    
    global InTable
    InTable=InTable+1
    ColumnCount=0
    TableStarted=[]  # Hack used as flag to show the table core elements were created
    
    try:
        for childnode in InputNode.childNodes:
            if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                # 0.29 adding support for <table:table-header-columns>
                if childnode.tagName=="table:table-header-columns":
                    for grandchildnode in childnode.childNodes:
                        if grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                            if grandchildnode.tagName=="table:table-column":
                            # if table:number-columns-repeated is present, add this number to ColumnCount
                            # if it is not present or not a number, just count the columns 
                                try:
                                    ColumnCount=ColumnCount+int(grandchildnode.getAttribute("table:number-columns-repeated"))
                                except ValueError:
                                    ColumnCount=ColumnCount+1
                        
                if childnode.tagName=="table:table-column":
                    # if table:number-columns-repeated is present, add this number to ColumnCount
                    # if it is not present or not a number, just count the columns 
                    try:
                        ColumnCount=ColumnCount+int(childnode.getAttribute("table:number-columns-repeated"))
                    except ValueError:
                        ColumnCount=ColumnCount+1
                elif childnode.tagName=="table:table-row":
                    TBodyNode=ProcessRow(childnode)
                elif childnode.tagName=="table:table-header-rows":
                    for gcnode in childnode.childNodes:
                        if gcnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                            if gcnode.tagName=="table:table-row":
                                TBodyNode=ProcessRow(gcnode)
    finally:
        InTable=InTable-1
                    
                    

# Process a list
# InputNode is assumed to be a <text:list> element
def ProcessListNode(InputNode,OutputNode):
    ColumnCount=0
    ListStarted=False  # Flag to show the list element (ol/ul) was created

    # Determine the list style name
    # If it is not in the current tag get it from the stack
    # If it IS in teh current tag, add it to the stack
    # Also sets a flag whether to pop the stack at the end
    # Note: stack is initialized with one member (empty string) so we don't check for empty stack
    global ListStyleNamesStack
    ListStyleName=InputNode.getAttribute("text:style-name")
    if ListStyleName=="":
        ListStyleName=ListStyleNamesStack[len(ListStyleNamesStack)-1]
        StyleNameAppended=False
    else:
        ListStyleNamesStack.append(ListStyleName)
        StyleNameAppended=True

    # Determine the list type string from the list style name
    # First we find the dictionary of levels (if the style is undefined it's empty)
    try:
        ListLevels=ListStylesDictionary[ListStyleName]
    except KeyError:
        #print ListStylesDictionary
        debug(3,"list style dictionary not found for: "+ListStyleName)
        ListLevels={}
    # Now implement the spec: if current level not found try lower levels
    ListType=""
    worklistlevel=CurrentListLevel
    while ListType=="" and worklistlevel>0:
        try:
            ListType=ListLevels[worklistlevel]
        except KeyError:
            pass
        worklistlevel=worklistlevel-1

    # If type is still not found, default to bullet list
    if ListType=="":
        debug(3,"list style not found for: "+ListStyleName)
        ListType="bullet"
        
    # now process child nodes of the list, which are, by spec, text:list-header and text:list-item
    for childnode in InputNode.childNodes:
        if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
            if childnode.tagName=="text:list-header":
                # process as normal text content, right into the parent node
                # This is because DITA has no list headers by its spec but ODT does by its spec
                ProcessTextNode(childnode,False,OutputNode)
            elif childnode.tagName=="text:list-item":
                # List data has started. Create list elements if not created before
                if not ListStarted:
                    ListStarted=True

                    if ListType=="number":
                        ListNode=InitialOutputDOM.createElement("ol")
                    else:
                        ListNode=InitialOutputDOM.createElement("ul")
                    OutputNode.appendChild(ListNode)
                    
                
                # Now create and fill the list item node
                ItemNode=InitialOutputDOM.createElement("li")
                ListNode.appendChild(ItemNode)
                ProcessTextNode(childnode,False,ItemNode)

    # Finally, pop the list styles stack if needed
    if StyleNameAppended:
        ListStyleNamesStack.pop()
        


# Process a text node
# if MainText then this is the main text node, <office:text>, or a <text:section> directly inside it
# Here, output is always to InitialOutputBody and headings are processed with new files
# mainText=False and OutputNode used for cells/items and other nested cases
# 0.38: added isFormula. It is set if <svg:desc>formula</svg:desc> was found and then <draw:> is being processed
def ProcessTextNode(InputNode,MainText=True,OutputNode=None,isFormula=False):

    foundMathML=False   # 0.38 this flag is used if isFormula is True and we already found <draw:object>,
                        # to exclude processing of <draw:image> after it



    
    if InputNode.nodeType<>xml.dom.Node.ELEMENT_NODE:
        debug(1,"ProcessStyleNode: non-element node received")
        # we can not process a non-element node so in this case we only output a debug message
        return()

    if MainText:
        OutputNode=InitialOutputBody

    for childnode in InputNode.childNodes:
        if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
            if (childnode.tagName=="text:p") or ((childnode.tagName=="text:h") and (InTable>0)):
                # process paragraph text
                attr_stylename=childnode.getAttribute("text:style-name")
#                try:
#                    outlevel=StylesOutlineLevels[attr_stylename]
#                except KeyError:
#                    outlevel=0

#                if attr_stylename.find("Head")>-1:
#                    print attr_stylename+" "+str(outlevel)
                
                
                
                # 0.022: Frame Mode support for headingN names
                if FrameMode and (attr_stylename in FrameModeStylesDictionary) and (InTable==0):
                    # Process a Frame mode heading - create new title for later breakup into Output DOM
                    # Save level into an attribute of the temp:topic node
                    debug(3,"Frame Mode creating topic, style name: "+attr_stylename)

                    NewLevel=FrameModeStylesDictionary[attr_stylename]

                    # 0.31: the temp topic is always created under the main output body 
                    TTNode=CreateNewTempTopic(InitialOutputBody)
                    TTNode.setAttribute("level",str(NewLevel))
                    
                    # process the title
                    ProcessParagraph(childnode,CurrentOutputTitle,set([]),True)
                    MoveId(CurrentOutputTitle,TTNode)
                    
#                # 0.027: outline level support (COMMENTED OUT)
#                elif (outlevel>0) and (InTable==0):
#                    print "Triggered"
#                    TTNode=CreateNewTempTopic(OutputNode)
#                    TTNode.setAttribute("level",str(outlevel))
#                    
#                    # process the title
#                    ProcessParagraph(childnode,CurrentOutputTitle,set([]),True)
                    
                
                else:
                    # process normal paragraph
                    try:
                        styleproperties=StylesDictionary[attr_stylename]
                    except KeyError: styleproperties=set([]) # 0.40 added handling "style not defined" here
                    paragraphnode=InitialOutputDOM.createElement("p")
                    OutputNode.appendChild(paragraphnode)
                    ProcessParagraph(childnode,paragraphnode,styleproperties)

                    # After a paragraph was created - even if it was nested in another paragraph -
                    # we need to process white space as in the start of a paragraph
                    global IsStartParagraph
                    IsStartParagraph=True
            elif (childnode.tagName=="text:h") and (InTable==0):
                # Process a heading - create new title for later breakup into Output DOM
                # Save level into an attribute of the temp:topic node


                #print "HEADER FOUND"

                try:
                    NewLevel=int(childnode.getAttribute("text:outline-level"))
                except ValueError:
                    NewLevel=1

                TTNode=CreateNewTempTopic(InitialOutputBody)
                TTNode.setAttribute("level",str(NewLevel))
                # pre-0.15: CreateOutputDOM(NewLevel)

                # process the title
                ProcessParagraph(childnode,CurrentOutputTitle,set([]),True)
                MoveId(CurrentOutputTitle,TTNode)
                    

                # pre-0.15: OutputNode=InitialOutputBody
                
                    
            elif childnode.tagName=="draw:image": # process an image by creating an image tag
                # 0.38 bypass image right after the MathML for a formula
                if not foundMathML:
                    imagenode=InitialOutputDOM.createElement("image")
                    image_href=childnode.getAttribute("xlink:href")
                    (image_root,image_ext)=os.path.splitext(image_href)
                    if not (image_ext.lower() in [".jpg",".png"]):
                         image_ext=".png"
                    imagenode.setAttribute("href",image_root+image_ext)
                    imagenode.setAttribute("placement","break")
                    OutputNode.appendChild(imagenode)
                    debug(3,"Image original "+image_href+" set as"+image_root+image_ext)
            elif (childnode.tagName=="draw:object" and (isFormula or AggressiveFormula)): # 0.38 process the MathMLobject in a formula
                fname=childnode.getAttribute("xlink:href")
                if fname[:2]=="./":
                    fname=fname[2:]
                fname=fname+"/content.xml"
                ToUnzip.add(fname)  # we will need to extract this MathML file
                foundMathML=True

                # now create the nodes
                eqnode=InitialOutputDOM.createElement("equation-inline")
                OutputNode.appendChild(eqnode)
                mathmlnode=InitialOutputDOM.createElement("mathml")
                eqnode.appendChild(mathmlnode)
                mathmlrefnode=InitialOutputDOM.createElement("mathmlref")
                mathmlrefnode.setAttribute("href",fname)
                mathmlnode.appendChild(mathmlrefnode)                
            elif childnode.tagName.find("draw:")>-1: # draws other than image - process content as text node
                # 0.38 check for formula
                foundFormula=False
                for grandchildnode in childnode.childNodes:
                    if grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                        if grandchildnode.tagName=="svg:desc":
                            descnode=grandchildnode.firstChild
                            if descnode.nodeType==xml.dom.Node.TEXT_NODE:
                                if descnode.data.strip().lower()=="formula":
                                    foundFormula=True

                ProcessTextNode(childnode,False,OutputNode,foundFormula)
            elif childnode.tagName=="table:table": # process table
                ProcessTableNode(childnode,OutputNode)
            elif childnode.tagName=="text:list":  # process list
                global CurrentListLevel
                CurrentListLevel+=1
                ProcessListNode(childnode,OutputNode)
                CurrentListLevel-=1
            elif childnode.tagName=="text:section":  # process section - exactly the same as if it was not there
                ProcessTextNode(childnode,MainText,OutputNode)
                       
            else:
                unprocessed(InputNode.tagName,childnode.tagName+" in ProcessTextNode")
    
# == POSTPROCESSING ==============================================

# Service for postprocessing
# Determine if a DITA node, presumed to be a paragraph-like entity (<p> or like it)
# contains only stuff allowed in a "plain" paragraph, without images or anything, thus suitable
# for a title, codeblock, etc
# always true for non-Element nodes (this simplifies recursion)
# IMPORTANT: relies on <p>s not being nested, so only for use after PostprocessFixNestedParagraphs

def IsSimpleParagraph(node):
    result=True
    if node.nodeType==xml.dom.Node.ELEMENT_NODE:
        if node.tagName in ["p","codeph","b","i"]:  # the usual formatting in a simple paragraph
            # check all child nodes
            for childnode in node.childNodes:
                result=(result and IsSimpleParagraph(childnode))
        else: # a different tag found - not a simple paragraph
            result=False
    return result

# Service: move id attribute from InputNode to OutputNode
# If InputNode has no id, do nothing
# if OutputNode already has an ID, replace the link in the table
# remove the ID attribute from InputNode
def MoveId(InputNode,OutputNode):
    global BookmarksDictionary
    if InputNode.hasAttribute("id"):
        IDToMove=InputNode.getAttribute("id")
        if OutputNode.hasAttribute("id") and OutputNode.getAttribute("id")<>IDToMove:
            NewID=OutputNode.getAttribute("id")

            # add to BookmarksDuplicate
            BookmarksDuplicate[IDToMove]=NewID
            #print "duplicate added at para move: "+BookmarksDuplicate[IDToMove]+" for ID "+IDToMove
                        

            # If any entry in BookmarksDuplicate points to this ID, repoint it
            for BookmarkID in BookmarksDuplicate:
                if BookmarksDuplicate[BookmarkID]==IDToMove:
                    BookmarksDuplicate[BookmarkID]=NewID
                    #print "duplicate changed at para move: "+BookmarksDuplicate[IDToMove]+" for ID "+IDToMove
            


        else:
            OutputNode.setAttribute("id",IDToMove)
        InputNode.removeAttribute("id")

# Save the "id" of a node which is to be removed
# In case of failure (no next and previous element sibling,and the body element is the parent) output a warning and break link
# If the node is not an element or has no ID, do nothing
def SaveId(node):
    if node.nodeType<>xml.dom.Node.ELEMENT_NODE:
        return
    if not node.hasAttribute("id"):
        return
    IDMoved=False
    # First, check if the next sibling exists and is an element...
    if node.nextSibling<>None:
        if node.nextSibling.nodeType==xml.dom.Node.ELEMENT_NODE:
            MoveId(node,node.nextSibling)
            IDMoved=True
    if not IDMoved:
        # next, try the previous sibling...
        if node.previousSibling<>None:
            if node.previousSibling.nodeType==xml.dom.Node.ELEMENT_NODE:
                MoveId(node,node.previousSibling)
                IDMoved=True
    if not IDMoved:
        # next, try the parent...
        if not (node.parentNode.tagName in ["conbody","section","context"]):
            MoveId(node,node.parentNode)
            IDMoved=True
    if not IDMoved:
        FailedID=node.getAttribute("id")
        debug(1,"Failed to move id: "+FailedID)
        # break link via BookmarksDuplicate
        BookmarksDuplicate[FailedID]=None

        # If any entry in BookmarksDuplicate points to this ID, repoint it
        for BookmarkID in BookmarksDuplicate:
            if BookmarksDuplicate[BookmarkID]==FailedID:
                BookmarksDuplicate[BookmarkID]=None    

        

# Service: remove all non-alpha characters from beginning of text under a DITA node
def RemoveAllBeforeAlpha(node):
    if node.nodeType<>xml.dom.Node.ELEMENT_NODE:
        debug(1,"RemoveAllBeforeAlpha called on non-element node!")
        return
    firsttext=FindFirstText(node)
    NeedMoreWork=True
    while (firsttext<>None) and NeedMoreWork:
        try:
            while not firsttext.data[0].isalpha():
                firsttext.data=firsttext.data[1:]
        except IndexError:pass
        if firsttext.data=='':
            DestroyNode(firsttext)
            firsttext=FindFirstText(node)
        else:
            NeedMoreWork=False

# Service: get the text content of a node, as one string
def GetTextAsString(node):
    if node.nodeType==xml.dom.Node.TEXT_NODE:
        return node.data
    elif node.nodeType==xml.dom.Node.ELEMENT_NODE:
        result=""
        for childnode in node.childNodes:
            result=result+GetTextAsString(childnode)
        return result

# Service: do a replace in the text content of a node
# FIXME: does not handle cases when the replace would straddle text nodes
def ReplaceText(node,old,new):
    if node.nodeType==xml.dom.Node.TEXT_NODE:
        node.data=node.data.replace(old,new)
    elif node.nodeType in [xml.dom.Node.ELEMENT_NODE,xml.dom.Node.DOCUMENT_NODE]:
        for childnode in node.childNodes:
            ReplaceText(childnode,old,new)
    
# Service: elevate an element node so its parent is within a set of tag names
# (or is the Document - flag this as an error)
# NOTE: we have to pass the DOM to this as well, because the DOM is used to create new elements
def ElevateNode(node,AllowedParentTagNames,DOM):
    # move node up by breaking up the previous one, one step at a time
    # for example, it was <node1>11<node2>22<p>text</p>22</node2>11</node1>
    # then it becomes: <node1>11<node2>22</node2><p>text</p><node2>22</node2>11</node1>
    # and finally: <node1>11<node2>22</node2></node1><p>text</p><node1><node2>22</node2>11</node1>
    # if <p> is the sole child of a parent node, that parent node is simply destroyed,
    #  with <p> taking its place
    try:
        while not (node.parentNode.tagName in AllowedParentTagNames):
            parent=node.parentNode
            newparent=parent.parentNode
            parenttag=parent.tagName
            # As the parent will be destroyed, move its ID to its first child - which may or may not be
            # the node being elevated
            MoveId(parent,parent.firstChild)

            if node.previousSibling<>None:
                newprevious=DOM.createElement(parenttag)
                newparent.insertBefore(newprevious,parent)
                CopyAttributes(parent,newprevious)
                while node.previousSibling<>None:
                    newprevious.appendChild(parent.firstChild)

            if node.nextSibling<>None:
                newnext=DOM.createElement(parenttag)
                newparent.insertBefore(newnext,parent.nextSibling)
                CopyAttributes(parent,newnext)
                while node.nextSibling<>None:
                    newnext.appendChild(node.nextSibling)

            newparent.insertBefore(node,parent)
            newparent.removeChild(parent)
            parent.unlink()
    except AttributeError:
        debug(1,"ElevateNode: node apparently raised to document - report error")
    

    



# Break paragraphs into two where a <temp:linebreak> is encountered
def PostprocessLinebreak(DOM):
    LNodes=DOM.getElementsByTagName("temp:linebreak")
    for node in LNodes:
        # Move the line break all the way up into the <p>, by breaking up nodes below the <p>
        # IMPORTANT: this will lead to an error if <p> is not among the parents
        # but we assume it is

        ElevateNode(node,["p"],DOM)
        

        # Now create a new paragraph before the current one, and move all nodes before the linebreak to it
        CurrentParagraph=node.parentNode  # This is a <p> because of the previous loop
        NewPrevParagraph=DOM.createElement("p")
        MoveId(CurrentParagraph,NewPrevParagraph)
        CurrentParagraph.parentNode.insertBefore(NewPrevParagraph,CurrentParagraph)
        CopyAttributes(CurrentParagraph,NewPrevParagraph)
        while node.previousSibling<>None:
            NewPrevParagraph.appendChild(node.parentNode.firstChild)

    # Now that the work has been done, destroy the <temp:linebreak> nodes
    for node in LNodes:
        DestroyNode(node)
            
        
# Elevate all temp:topic elements into the body - no longer needed in 0.31 
#def PostprocessElevateTempTopic(DOM):
#    TNodes=DOM.getElementsByTagName("temp:topic")
#    for node in TNodes:
#        ElevateNode(node,"conbody",DOM)
    


# Fix <p> tags inside tags that can't contain them
# This postprocessing step is required because of the way the main processing works
# Perhaps we need to do this for some other tags? For <p> it's certain
# First, a piece of data - which tags can contain <p>
# WARNING: This is from DITA help and IS PROBABLY TOO MUCH
# If illicit <p>s show up, the tag om which they show up must be removed from here
CanContainP=["dita","data","title","shortdesc","desc","note","lq","q","sli",
             "li","itemgroup","dthd","ddhd","dt","dd","fig","figgroup","pre",
             "lines","ph","alt","object","stentry","draft-comment","fn",
             "indexterm","index-base","cite","xref","table","tgroup","entry",
             "author","source","publisher","copyright","copyrholder","category",
             "keywords","prodinfo","prodname","brand","series","platform","prognum",
             "featnum","component","topic","navtitle","searchtitle","abstract",
             "body","section","example","prolog","metadata","related-links",
             "linktext","linkinfo","linkpool","concept","conbody","task","taskbody",
             "prereq","context","steps","steps-unordered","step","cmd","info",
             "substeps","substep","tutorialinfo","stepxmp","choice","chhead",
             "choptionhd","chdeschd","chrow","choption","chdesc","stepresult","result",
             "postreq","reference","refbody","refsyn","properties","prophead",
             "proptypehd","propvaluehd","propdeschd","property","proptype","propvalue",
             "propdesc","glossentry","glossterm","glossdef","uicontrol","screen",
             "tt","sup","sub","codeph","codeblock","var","synph","oper","delim","sep",
             "parml","plentry","pt","pd","syntaxdiagram","synblk","groupseq",
             "groupchoice","groupcomp","fragment","fragref","synnote","repsep",
             "msgph","msgblock","filepath","userinput","systemoutput","area","coords",
             "index-see","index-see-also","index-sort-as"]
def PostprocessFixNestedParagraphs(DOM):
    # the obvious and wrong solution is a simple scan of DOM.getElementsByTagName()
    # but every raise of a <p> disrupts node structure and might delete any of the found nodes
    # so we have to scan anew each time, until we find no illicitly located <p>s

    SearchAgain=True

    while SearchAgain:
        SearchAgain=False
        PNodes=DOM.getElementsByTagName("p")
        for node in PNodes:
            if not (node.parentNode.tagName in CanContainP):
                SearchAgain=True
                ElevateNode(node,CanContainP,DOM)
                break



# Move blank space only <b> and <i> and <note> text into parent tags and destroy b/i/note (but not codeph) 
# Then, blank space only paragraphs into empty paragraphs
# We don't remove them at this stage as they may be needed, notably for code blocks
def PostprocessSpace(DOM):

    # Internal text to see if there is anything except whitespace in a node, using recursion
    # IMPORTANT: relies on all nodes being <p> or formatting, so only used after IsSimpleParagraph check
    def IsBlank(node):
        if node.nodeType==xml.dom.Node.TEXT_NODE:
            return (node.data.isspace() or (node.data==''))
        elif node.nodeType==xml.dom.Node.ELEMENT_NODE:
            Blank=True
            for childnode in node.childNodes:
                Blank=(Blank and IsBlank(childnode))
            return Blank
        else: return True
        
    # move whitespace only <b>/<i>/<note> into parent
    # note: has to be done repeatedly until no hits are found, because they might be nested
    # only works if there are no children except one text - this is on purpose
    NeedMoreWork=True
    while NeedMoreWork:
        NeedMoreWork=False
        NodesToDelete=[]
        for tag in ["b","i","note"]:
            TNodes=DOM.getElementsByTagName(tag)
            for TNode in TNodes:
                if TNode.hasChildNodes():
                    if TNode.firstChild==TNode.lastChild:
                        if TNode.firstChild.nodeType==xml.dom.Node.TEXT_NODE:
                            if TNode.firstChild.data.isspace():
                                # found a space-only child
                                TNode.parentNode.insertBefore(TNode.firstChild,TNode)
                                NodesToDelete.append(TNode)
                                NeedMoreWork=True
        for node in NodesToDelete:
            DestroyNode(node)
                            
    
    # handle blank paragraphs
    PNodes=DOM.getElementsByTagName("p")
    for node in PNodes:
        if IsSimpleParagraph(node):
            if IsBlank(node):
                # remove all children but leave node intact
                for childnode in node.childNodes:
                    DestroyNode(childnode)
        
    

    
# Join immediately adjacent <ul> lists
def PostprocessJoinLists(DOM):
    # initialize list of nodes to delete
    # make it a set, just to avoid duplicates
    NodesToDelete=set([])

    ulNodes=DOM.getElementsByTagName("ul")
    for node in ulNodes:
        if node.previousSibling<>None:
            if node.previousSibling.nodeType==xml.dom.Node.ELEMENT_NODE:
                if node.previousSibling.tagName=="ul":
                    # the previous sibling is another list - need to merge
                    # but first we need to check if this is not set for deletion,
                    # and if it is move to previous siblings until finding a list not set for deletion
                    # this allows normal merging of 3+ lists without repeated iterations
                    destination=node.previousSibling
                    while destination in NodesToDelete:
                        destination=destination.previousSibling
                        # if the new destination is not a <ul> something is broken in this code
                        # (because only nodes after a <ul> can be flagged for deletion)
                        # report error and quit function immediately
                        if destination==None:
                            debug(0,"Internal error in PostprocessJoinLists - no previous sibling")
                            return()
                        if not destination.nodeType==xml.dom.Node.ELEMENT_NODE:
                            debug(0,"Internal error in PostprocessJoinLists - previous sibling not element")
                            return()
                        if not destination.tagName=="ul":
                            debug(0,"Internal error in PostprocessJoinLists - previous sibling not <ul>")
                            return()

                    # move all child nodes and ID to destination, and flad the node for deletion
                    MoveChildNodes(node,destination)
                    MoveId(node,destination)
                    NodesToDelete.add(node)

    # delete the flagged nodes
    for node in NodesToDelete:
        node.parentNode.removeChild(node)
        node.unlink()

# 0.29 Fix nexted lists where each nesting only has one <li> and a list in there -
# an artefact of ODT nested list processsing
# Intentionally called *after* PostprocessJoinLists to avoid undesirable joining
def PostprocessNestedOneLists(DOM):
    listtags=["ul","ol"]
    for nowtag in listtags:
        foundlists=DOM.getElementsByTagName(nowtag)
        MoreWork=True
        while MoreWork:
            Triggered=False
            for listnode in foundlists:
                if listnode.hasChildNodes():
                    if listnode.firstChild==listnode.lastChild:
                        childnode=listnode.firstChild
                        if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                            if childnode.tagName=="li": # so far we have only one child and it is <li>
                                if childnode.hasChildNodes():
                                    if childnode.firstChild==childnode.lastChild:
                                        grandchildnode=childnode.firstChild
                                        if grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                                                if grandchildnode.tagName in listtags:
                                                    # we have the specific situation
                                                    # lift the newfound list before the list node,
                                                    # remove the list node and start again
                                                    listnode.parentNode.insertBefore(grandchildnode,listnode)
                                                    DestroyNode(listnode)
                                                    Triggered=True
                                                    break # will apply to the for loop
            MoreWork=Triggered

                                                    



# Join immediately adjacent tags. Called for codeph, u and i . Made from the one for ul's, but FIXME no time to investigate joining them
def PostprocessJoinTags(DOM,TagName):
    # initialize list of nodes to delete
    # make it a set, just to avoid duplicates
    NodesToDelete=set([])

    ulNodes=DOM.getElementsByTagName(TagName)
    for node in ulNodes:
        if node.previousSibling<>None:
            if node.previousSibling.nodeType==xml.dom.Node.ELEMENT_NODE:
                if node.previousSibling.tagName==TagName:
                    # the previous sibling is another of the same tag - need to merge
                    # but first we need to check if this is not set for deletion,
                    # and if it is move to previous siblings until finding a list not set for deletion
                    # this allows normal merging of 3+ tags without repeated iterations
                    destination=node.previousSibling
                    while destination in NodesToDelete:
                        destination=destination.previousSibling
                        # if the new destination is not a <ul> something is broken in this code
                        # (because only nodes after a <ul> can be flagged for deletion)
                        # report error and quit function immediately
                        if destination==None:
                            debug(0,"Internal error in PostprocessJoinTags - no previous sibling")
                            return()
                        if not destination.nodeType==xml.dom.Node.ELEMENT_NODE:
                            debug(0,"Internal error in PostprocessJoinTags - previous sibling not element")
                            return()
                        if not destination.tagName==TagName:
                            debug(0,"Internal error in PostprocessJoinTagss - previous sibling not same tag")
                            return()

                    # move all child nodes and ID to destination, and flad the node for deletion
                    MoveChildNodes(node,destination)
                    MoveId(node,destination)
                    NodesToDelete.add(node)

    # delete the flagged nodes
    for node in NodesToDelete:
        node.parentNode.removeChild(node)
        node.unlink()
    


# Remove first <p> (and move its attributes and content outside it) for several element types
def PostprocessFirstP(DOM):
    for tagname in ["li","entry","fn"]: # FIXME: are there more?
        TNodes=DOM.getElementsByTagName(tagname)
        for node in TNodes:
            if node.hasChildNodes():
                if node.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE:
                    if node.firstChild.tagName=="p":
                        PNode=node.firstChild
                        MoveId(PNode,node)
                        CopyAttributes(PNode,node)
                        while PNode.hasChildNodes():
                            node.insertBefore(PNode.firstChild,PNode)
                        node.removeChild(PNode)
                        PNode.unlink()

# Remove <note> tags from footnotes
# The first one goes into <fn> directly, the rest become <p>
def PostprocessFootnotes(DOM):
        TNodes=DOM.getElementsByTagName("fn")
        for node in TNodes:
            if node.hasChildNodes():
                # process the first <note> tag
                if node.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE:
                    if node.firstChild.tagName=="note":
                        PNode=node.firstChild
                        MoveId(PNode,node)
                        CopyAttributes(PNode,node)
                        while PNode.hasChildNodes():
                            node.insertBefore(PNode.firstChild,PNode)
                        node.removeChild(PNode)
                        PNode.unlink()

                # process the rest now
                for childnode in node.childNodes:
                    if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                        if childnode.tagName=="note":
                            childnode.tagName="p"
                            try:
                                childnode.removeAttribute("type")
                            except NotFoundErr: pass
                                                    
# Postprocess the tables
def PostprocessTables(DOM):

    # Internal routine for moving stuff out of a table entry/cell and into the text
    # Everything until a <p>, if it exists, is moved to a new <p>
    # TableNode must be a child of OutputNode (it is the node before which text is to be put)
    # Also, if a new <p> is created, the otherprops attribute from InputNode is copied to it
    #  and the id attribute if present is moved
    def MoveFromEntry(InputNode,OutputNode,TableNode):
        if InputNode.hasChildNodes(): # only do anything at all if there are nodes to move
            # Determine if a first <p> is needed
            NeedToCreateP=True
            if InputNode.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE:
                if InputNode.firstChild.tagName=="p":
                    NeedToCreateP=False


            if NeedToCreateP:
                # Create the <p>
                PNode=DOM.createElement("p")
                OutputNode.insertBefore(PNode,TableNode)

                # Copy otherprops attribute
                if InputNode.hasAttribute("otherprops"):
                    PNode.setAttribute("otherprops",InputNode.getAttribute("otherprops"))

                # Move id attribute
                MoveId(InputNode,OutputNode)

                
                # Move nodes to the <p>
                MoveNext=True
                while MoveNext:
                    PNode.appendChild(InputNode.firstChild)
                    if not InputNode.hasChildNodes():
                        MoveNext=False # stop the moving if we are out of nodes
                    elif InputNode.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE:
                        if InputNode.firstChild.tagName=="p":
                            MoveNext=False  # stop the moving if the next node is a <p>

            while InputNode.hasChildNodes(): # move remaining nodes, all assumed to be <p>
                OutputNode.insertBefore(InputNode.firstChild,TableNode)
                
                    
        
    

    
    NodesToDelete=set([])
   
    tableNodes=DOM.getElementsByTagName("table")
    
    for tablenode in tableNodes:
        # get the TGroup node - it is the first child of the table as of now
        # FIXME should I rely on this?
        TGroupNode=tablenode.firstChild

        NumberOfColumns=int(TGroupNode.getAttribute("cols"))

        # A table can only have one title. Flag if found
        TitleFound=False
        # If a caption exists immediately before, use it
        if tablenode.previousSibling<>None:
            if tablenode.previousSibling.nodeType==xml.dom.Node.ELEMENT_NODE:
                if tablenode.previousSibling.tagName=="p":
                    if tablenode.previousSibling.getAttribute("otherprops").find("caption")>-1:
                        captionnode=tablenode.previousSibling  
                        if captionnode.hasChildNodes() and IsSimpleParagraph(captionnode):
                            # only bother if the caption is not empty and content is suitable for title 
                            # We need to get around any paragraph-wide child elements
                            captionsourcenode=captionnode
                            while captionsourcenode.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE and \
                                captionsourcenode.firstChild==captionsourcenode.lastChild:
                                    captionsourcenode=captionsourcenode.firstChild
                            if captionsourcenode.hasChildNodes(): # only bother if the caption
                                                                # AFTER all paragraph-wide tags is not empty
                                # move the caption into the table title
                                # FIXME: add removal of Table N
                                tabletitlenode=DOM.createElement("title")
                                tablenode.insertBefore(tabletitlenode,tablenode.firstChild)
                                MoveChildNodes(captionsourcenode,tabletitlenode)
                                TitleFound=True

                                # copy any attributes from caption to table itself
                                CopyAttributes(captionnode,tablenode)
                                
                                # carefully delete the caption and all para-wide child nodes we went into
                                DestroyNode(captionnode)
                                    

                                
        
        
        
        # find the <tbody> and within it the first <row> and <entry>, and column names
        # FIXME? this relies on exactly one tgroup in table (very first child)
        #  and one <tbody> in table and
        #  its first child being <row>
        # *and* on the row's first child being <entry>
        FirstColumn=""
        LastColumn=""
        TBodyNode=None
        for childnode in TGroupNode.childNodes:
            if childnode.nodeType==xml.dom.Node.ELEMENT_NODE:
                #print childnode.tagName
                if childnode.tagName=="colspec":
                    colname=childnode.getAttribute("colname")
                    # assign first column if it's first - not assigned, and last column in all occasions
                    if FirstColumn=="":
                        FirstColumn=colname
                    LastColumn=colname
                elif childnode.tagName=="tbody":
                    TBodyNode=childnode

        # 0.30 body not found - report error, remove table
        if TBodyNode==None:
            debug(1,"ATTENTION: table without tbody element removed")
            NodesToDelete.add(tablenode)
            continue
        
        if FirstColumn=="":
            debug(1,"PostprocessTables: failed to determine first column, wrong results likely")
        FirstRowNode=TBodyNode.firstChild
        FirstEntryNode=FirstRowNode.firstChild

        # Check if the first row spans all the table (or there is only one column)
        # These are handled in a special way - being independent paragraphs if not a header
        # or the table title if, indeed, a header and if no caption is present
        # FIXME: for now, if there is a simple non-headered table with 1 column, this moves row 1 out of it
        #  this might not be forrect in all cases
        if (FirstEntryNode.getAttribute("namest")==FirstColumn and \
           FirstEntryNode.getAttribute("nameend")==LastColumn) or NumberOfColumns==1:
                if FirstEntryNode.getAttribute("otherprops").find("header")>-1:
                   # this is a header - use as caption if title not found yet, otherwise don't change
                   if not TitleFound:
                        captionsourcenode=FirstEntryNode
                        # get through paragraph-wide tags
                        while captionsourcenode.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE and \
                            captionsourcenode.firstChild==captionsourcenode.lastChild:
                                captionsourcenode=captionsourcenode.firstChild

                        # move the caption into the table title
                        # FIXME: add removal of Table N
                        tabletitlenode=DOM.createElement("title")
                        tablenode.insertBefore(tabletitlenode,tablenode.firstChild)
                        MoveChildNodes(captionsourcenode,tabletitlenode)

                        # copy any attributes from caption to table itself
                        CopyAttributes(FirstRowNode,tablenode)

                        TitleFound=True

                        # Remove the entire first row node 
                        DestroyNode(FirstRowNode)
                else:
                    # this is not a header, make it independent paragraphs
                    # the child nodes are to be moved immediately before the table node
                    MoveFromEntry(FirstEntryNode,tablenode.parentNode,tablenode)

                    # Remove the entire first row node
                    DestroyNode(FirstRowNode)
        
        # Now, recalculate FirstRowNode 
        FirstRowNode=TBodyNode.firstChild

        # If no more rows do remain, flag table for deletion
        if FirstRowNode==None:
            NodesToDelete.add(tablenode)
        else:
            # continue processing if nodes remain
            # now, if any rows in the beginning have the header in otherprops on ALL entries
            #  then make them header rows
            # BUT avoid doing this if only one row remains
            HeaderAdded=False
            HeaderRowFound=True
            while (FirstRowNode.nextSibling<>None) and HeaderRowFound:
                HeaderRowFound=True  # really redundant but kept for clarity
                for EntryNode in (FirstRowNode.childNodes): # assume all row child nodes are <entry>
                    if EntryNode.getAttribute("otherprops").find("header")==-1:
                        HeaderRowFound=False

                if HeaderRowFound:                        
                    # add the header if not previously added
                    if not HeaderAdded:
                        THeadNode=DOM.createElement("thead")
                        TGroupNode.insertBefore(THeadNode,TBodyNode)
                        HeaderAdded=True
                    # move the row to the header
                    THeadNode.appendChild(FirstRowNode)

                    # remove paragraph-wide <b> and <i> tags from every node in the new header row
                    for EntryNode in FirstRowNode.childNodes:
                        if EntryNode.hasChildNodes():
                            ToCheck=True
                            while ToCheck:
                                ToCheck=False
                                if not EntryNode.hasChildNodes(): # check added in 0.32
                                    break
                                if (EntryNode.firstChild==EntryNode.lastChild) and \
                                   EntryNode.firstChild.nodeType==xml.dom.Node.ELEMENT_NODE:
                                    if EntryNode.firstChild.tagName in ["b","i"]:
                                        # Will need to check again after this is complete
                                        ToCheck=True
                                        
                                        # Move all the information and remove the child node
                                        childtoremove=EntryNode.firstChild
                                        MoveChildNodes(childtoremove,EntryNode)
                                        DestroyNode(childtoremove)
                                    
                    

                    # recalculate FirstRowNode
                    FirstRowNode=TBodyNode.firstChild

            if TitleFound:
                # Remove "Table something" from title
                tnode=FindFirstText(tabletitlenode)
                if tnode<>None:
                    data=tnode.data

                    # Remove spaces from beginning - this will only get back into the note if an infostring is found
                    # If data gets empty anywhere in the process, never mind
                    # FIXME? This would not process titles where whitespace is the whole first text node
                    #  rather unlikely but still possible 
                    try:
                        while data[0].isspace():
                            data=data[1:]
                    except IndexError: pass

                    # Check for "Table"
                    # If the string is found, remove it and flag
                    StringFound=False
                    if data.lower().find("table")==0:
                        data=data[5:]
                        StringFound=True

                    if StringFound:
                        # If the string was found, remove any :, number etc (nonalpha chars)
                        #  immediately after it
                        # If data gets empty, that's ok
                        try:
                            while not data[0].isalpha():
                                data=data[1:]
                        except IndexError: pass

                        if data=='':
                            # remove the text node, and handle possible empty nodes
                            DestroyNode(tnode)
                            RemoveChildlessElementChildrenPara(tabletitlenode)
                        else:
                            # set text node to new data
                            tnode.data=data
                        RemoveAllBeforeAlpha(tabletitlenode)

    # delete flagged tables
    for node in NodesToDelete:
        DestroyNode(node)


# Postprocesing to detect code blocks
# For now, only paragraphs with a paragraph monospace property (or blank ones) are contemplated
# relies on previous postprocessing
def PostprocessCodeblock(DOM):
    # initialize list of nodes to delete
    # make it a set, just to avoid duplicates
    NodesToDelete=set([])

    PNodes=DOM.getElementsByTagName("p")

    for PNode in PNodes:
        if IsSimpleParagraph(PNode): # only check for anything at all if the paragraph has only formatting in it
                                     # not images, tables, etc
            # First we work out if this is a blank paragraph or a valid code paragraph
            IsCode=False
            if PNode.hasChildNodes():
                IsBlank=False
                if PNode.getAttribute("otherprops").find("codeph")>-1:
                    # Check if there is a paragraph-wide <codeph>, which may be under others
                    # If found, note and remove it
                    if PNode.firstChild==PNode.lastChild:
                        ToCheck=True
                        CheckNode=PNode.firstChild
                        while ToCheck:
                            if CheckNode.nodeType<>xml.dom.Node.ELEMENT_NODE:
                                ToCheck=False
                            elif CheckNode.tagName=="codeph":
                                CodephNode=CheckNode
                                IsCode=True
                                ToCheck=False
                            elif CheckNode.hasChildNodes():
                                ToCheck=(CheckNode.firstChild==CheckNode.lastChild)
                                CheckNode=CheckNode.firstChild
                            else:
                                ToCheck=False

                        if IsCode:
                            MoveChildNodes(CheckNode,CheckNode.parentNode)
                            DestroyNode(CheckNode)

            else:
                IsBlank=True

            # Check if the previous node, if any, is a codeblock
            PrevCodeblock=False
            PrevNode=PNode.previousSibling
            while (PrevNode<>None) and (PrevNode in NodesToDelete):
                PrevNode=PrevNode.previousSibling
            
            if PrevNode<>None:
                if PrevNode.nodeType==xml.dom.Node.ELEMENT_NODE:
                    if PrevNode.tagName=="codeblock":
                        PrevCodeblock=True

            if PrevCodeblock:
                if IsBlank or IsCode:
                    MoveChildNodes(PNode,PrevNode,"\n")
                    # This works for a blank just as well - the \n is added and then nothing is there to move

                    MoveId(PNode,PrevNode)
                    
                    NodesToDelete.add(PNode)
            elif IsCode:
                # Change paragraph into codeblock
                PNode.tagName="codeblock"

    # delete the flagged nodes
    for node in NodesToDelete:
        node.parentNode.removeChild(node)
        node.unlink()

    # Additional codeblock processing
    # Remove trailing newlines from all codeblocks
    # This could also be done without a search, by keeping a list of created codeblocks
    
    CodeblockNodes=DOM.getElementsByTagName("codeblock")

    for CNode in CodeblockNodes:
        if CNode.hasChildNodes():
            lastchild=CNode.lastChild
            if lastchild.nodeType==xml.dom.Node.TEXT_NODE:
                data=lastchild.data
                try:
                    while data<>"" and data[len(data)-1]=="\n":
                        data=data[:-1]
                except IndexError: pass
                if data=="":
                    DestroyNode(lastchild)
                else:
                    lastchild.data=data


# Postprocess notes
# Remove all supported note information strings, reflecting them in type attribute instead
NoteTypes=["note","attention","caution","danger","fastpath","important",
           "remember","restriction","tip"]

def PostprocessNotes(DOM):
    NoteNodes=DOM.getElementsByTagName("note")
    for NNode in NoteNodes:
        tnode=FindFirstText(NNode)
        if tnode<>None:
            data=tnode.data

            # Remove spaces from beginning - this will only get back into the note if an infostring is found
            # If data gets empty anywhere in the process, we skip this note entirely
            # FIXME? This would not process notes where whitespace is the whole first text node
            #  rather unlikely but still possible 
            try:
                while data[0].isspace():
                    data=data[1:]
            except IndexError: continue

            # Check for all note type info strings
            # If the string is found, remove it and flag
            StringFound=False
            for notetype in NoteTypes:
                if data.lower().find(notetype)==0:
                    data=data[len(notetype):]
                    NNode.setAttribute("type",notetype)
                    StringFound=True

            if StringFound:
                # If the string was found, remove any : etc (nonalphanum chars) immediately after it
                # If data gets empty, that's ok
                try:
                    while not data[0].isalnum():
                        data=data[1:]
                except IndexError: pass

                if data=='':
                    # remove the text node, and handle possible empty nodes
                    DestroyNode(tnode)
                    RemoveChildlessElementChildrenPara(NNode)
                else:
                    # set text node to new data
                    tnode.data=data

# Postprocess notes from paragraphs
# Convert all simple paragraphs satrting with supported note information strings to notes
NoteTypes=["note","attention","caution","danger","fastpath","important",
           "remember","restriction","tip"]

def PostprocessNotesFromParagraphs(DOM):
    PNodes=DOM.getElementsByTagName("p")
    for PNode in PNodes:
      if IsSimpleParagraph(PNode):
        tnode=FindFirstText(PNode)
        if tnode<>None:
            data=tnode.data

            # Remove spaces from beginning - this will only get back into the note if an infostring is found
            # If data gets empty anywhere in the process, we skip this note entirely
            # FIXME? This would not process notes where whitespace is the whole first text node
            #  rather unlikely but still possible 
            try:
                while data[0].isspace():
                    data=data[1:]
            except IndexError: continue

            # Check for all note type info strings
            # If the string is found, remove it and flag
            StringFound=False
            for notetype in NoteTypes:
                if data.lower().find(notetype+":")==0:
                    data=data[len(notetype):]
                    PNode.tagName="note"
                    PNode.setAttribute("type",notetype)
                    StringFound=True

            if StringFound:
                # If the string was found, remove any : etc (nonalphanum chars) immediately after it
                # If data gets empty, that's ok
                try:
                    while not data[0].isalnum():
                        data=data[1:]
                except IndexError: pass

                if data=='':
                    # remove the text node, and handle possible empty nodes
                    DestroyNode(tnode)
                    RemoveChildlessElementChildrenPara(PNode)
                else:
                    # set text node to new data
                    tnode.data=data



# Postprocess - break up the initial DOM to create output DOMs
# Also set Output DOM IDs and types accordingly
# In a title, [r] means reference and [t] means task
# With a task, stub steps are also created
# FIXME? Takes no arguments because it relies on InitialOutputBody
def PostprocessBreakupIntoTopics():

    # Internal recursion function - copying nodes, saving IDs in BookmarksDictionary
    # InputNode is copied as a child to OutputNode
    # CurrentOutputDOM is used (but not modified so namespace is not an issue)
    def CopyToOutput(InputNode,OutputNode):
        if InputNode.nodeType==xml.dom.Node.TEXT_NODE:
            textnode=CurrentOutputDOM.createTextNode(InputNode.data)
            OutputNode.appendChild(textnode)
        elif InputNode.nodeType==xml.dom.Node.ELEMENT_NODE:
            elemnode=CurrentOutputDOM.createElement(InputNode.tagName)
            OutputNode.appendChild(elemnode)
            #if InputNode.tagName=="xref":
                #print "Transferred xref: "+InputNode.getAttribute("href")
            CopyAttributes(InputNode,elemnode)

            #process bookmark if ID found
            BookmarkID=InputNode.getAttribute("id")

            
            if BookmarkID<>"":
                #print "ID found"
                if BookmarkID in AllIDs:
                    debug(1,"Duplicate bookmark ID: "+BookmarkID)
                else:
                    AllIDs.append(BookmarkID)
                    Ref="#"+CurrentDocID+"/"+BookmarkID
                    BookmarksDictionary[BookmarkID]=(CurrentDocID,Ref) 

            # process child nodes
            for childnode in InputNode.childNodes:
                CopyToOutput(childnode,elemnode)

    # MAIN FUNCTION CODE

    # get the MiniDOM implementation - required for creating new DOMs 
    impl=xml.dom.getDOMImplementation("minidom")


    for node in InitialOutputBody.childNodes:
        if node.nodeType==xml.dom.Node.ELEMENT_NODE:
            if node.tagName=="temp:topic":
                # Start a new topic

                # Get the level

                try:
                    level=int(node.getAttribute("level"))
                except ValueError:
                    debug(1,"No level set in temp:title")
                    level=1

                # Get the existing bookmark ID, if present
                BookmarkID=node.getAttribute("id")
                #print ("BookmarkID="+BookmarkID)


                # Process the title node - get doc title
                TitleNode=node.getElementsByTagName("title").item(0)
                doctitle=GetTextAsString(TitleNode).strip()

                
                # If this is an empty topic - blank title and next sibling is a <temp:topic> - move ID to next, and skip it
		# if the title is empty but the topic is not, change title to "notitle"
                if doctitle=="":
                    doctitle="notitle"
                    if node.nextSibling<>None:
                        if node.nextSibling.nodeType==xml.dom.Node.ELEMENT_NODE:
                            if node.nextSibling.tagName=="temp:topic":
                                MoveId(node,node.nextSibling)
                                continue

                # Determine document type (and fix title)

                doctypestr="concept" # this is the default
                if doctitle.find("[r]")>-1:
                    doctypestr="reference"
                    ReplaceText(TitleNode,"[r]","")
                    # remove any double space or final space this may have created
                    ReplaceText(TitleNode,"  "," ")
                    ltnode=FindLastText(TitleNode)
                    if ltnode.data[len(ltnode.data)-1].isspace():
                        ltnode.data=ltnode.data[:len(ltnode.data)-1]
                        
                if doctitle.find("[t]")>-1:
                    doctypestr="task"
                    ReplaceText(TitleNode,"[t]","")
                    # remove any double space or final space this may have created
                    ReplaceText(TitleNode,"  "," ")
                    ltnode=FindLastText(TitleNode)
                    if ltnode.data[len(ltnode.data)-1].isspace():
                        ltnode.data=ltnode.data[:len(ltnode.data)-1]

                if doctitle.find("[c]")>-1:
                    ReplaceText(TitleNode,"[c]","")
                    
                    
                # changed in 0.36 on user request
                # determine document (topic) ID - based on the title
                # [r] and [t] are removed by this time, so we get the title text anew
                CurrentDocID=GetTextAsString(TitleNode).strip().lower()

                # 0.42 if the ID is empty,s et it to "notitle"
                if CurrentDocID=="":
                    CurrentDocID="notitle"

                # replace non-alphanumeric characters with "_"
                for n in range(0,len(CurrentDocID)):
                    if not CurrentDocID[n].isalnum():
                        CurrentDocID=CurrentDocID[:n]+"_"+CurrentDocID[n+1:]

                # remove multiple "_" - replace with single "_"
                while CurrentDocID.find("__")>=0:
                    CurrentDocID=CurrentDocID.replace("__","_")

                # truncate
                CurrentDocID=CurrentDocID[:30]
                        
                        

                # prefix the ID with conc/task/refe
                if not DoNotPrefix:
                    if doctypestr=="reference":
                        CurrentDocID="r_"+CurrentDocID
                    elif doctypestr=="task":
                        CurrentDocID="t_"+CurrentDocID
                    else: # concept is the default
                       CurrentDocID="c_"+CurrentDocID
                else: # protection from leading number
                    if not CurrentDocID[0].isaplha():
                        CurrentDocID="a"+CurrentDocID


                #remove trailing _'s unless it makes the id too short
                while (len(CurrentDocID)>5) and (CurrentDocID[len(CurrentDocID)-1]=="_"):
                    CurrentDocID=CurrentDocID[:len(CurrentDocID)-1]


                # replace non-ascii characters with underscores
                for ii in range(0,len(CurrentDocID)-1):
                    try:
                        CurrentDocID[ii].decode("ascii")
                    except (UnicodeDecodeError,UnicodeEncodeError):
                        CurrentDocID=CurrentDocID[:ii]+"_"+CurrentDocID[ii+1:]
                   
                # make sure doc ID is unique
                IDbase=CurrentDocID
                n=1
                while CurrentDocID in AllIDs:
                    CurrentDocID=IDbase+str(n)
                    n=n+1

                AllIDs.append(CurrentDocID)
                #TEMP
                #print "Topic ID: "+CurrentDocID

#                if doctypestr=="reference":
#                    doctype=impl.createDocumentType("reference", "-//IBM//DTD DITA IBM Reference//EN", "ibm-reference.dtd")
#                elif doctypestr=="task":
#                    doctype=impl.createDocumentType("task", "-//IBM//DTD DITA IBM Task//EN", "ibm-task.dtd")
#                else: # concept is the default
#                    doctype=impl.createDocumentType("concept", "-//IBM//DTD DITA IBM Concept//EN", "ibm-concept.dtd")

                # OASIS
                doctype=impl.createDocumentType(doctypestr, "-//OASIS//DTD DITA Composite//EN", "ditabase.dtd")

                CurrentOutputDOM=impl.createDocument(None,doctypestr,doctype)

                # set topic "id" attribute to the newly created ID
                # and xml:lang as required - NOT NEEDED FOR OASIS
                TopicElement=GetElementChild(CurrentOutputDOM)
                TopicElement.setAttribute("id",CurrentDocID)
                # TopicElement.setAttribute("xml:lang","en-us")
                #OASIS seems to require xmlns:ditaarch
                TopicElement.setAttribute("xmlns:ditaarch","http://dita.oasis-open.org/architecture/2005/")
                

                # add the DOM to the list of output DOMs
                OutputDOMs.append((CurrentOutputDOM,level))
                
                # Process bookmark if required
                if BookmarkID<>"":
                    BookmarksDictionary[BookmarkID]=(CurrentDocID,"#"+CurrentDocID)

                # create the title and copy content from the old title node
                # NOTE: I do not use CopyToOutput for the entire title because I do not want to preserver the ID value or to process it in that routine
                
                NewTitleElement=CurrentOutputDOM.createElement("title")
                TopicElement.appendChild(NewTitleElement)
                for childnode in TitleNode.childNodes:
                    CopyToOutput(childnode,NewTitleElement)
                
                # create the empty shortdesc
                ShortdescElement=CurrentOutputDOM.createElement("shortdesc")
                TopicElement.appendChild(ShortdescElement)

                # Copy the prolog node only if it contains at least one indexterm 
                PrologNode=node.getElementsByTagName("title").item(0)
                if PrologNode.getElementsByTagName("indexterm").length>0:
                    CopyToOutput(PrologNode,TopicElement)

    

                # create the body, depending on document type
                if doctypestr=="reference":
                    # create new reference body, add the output body as its section
                    RefBodyNode=CurrentOutputDOM.createElement("refbody")
                    TopicElement.appendChild(RefBodyNode)
                    CurrentOutputBody=CurrentOutputDOM.createElement("section")
                    RefBodyNode.appendChild(CurrentOutputBody)                   

                    
                elif doctypestr=="task":
                    # create new task body, add the output body as its context, and add dummy steps
                    TaskBodyNode=CurrentOutputDOM.createElement("taskbody")
                    TopicElement.appendChild(TaskBodyNode)
                    CurrentOutputBody=CurrentOutputDOM.createElement("context")
                    TaskBodyNode.appendChild(CurrentOutputBody)                   

                    StepsNode=CurrentOutputDOM.createElement("steps")
                    TaskBodyNode.appendChild(StepsNode)
                    StepNode=CurrentOutputDOM.createElement("step")
                    StepsNode.appendChild(StepNode)
                    CmdNode=CurrentOutputDOM.createElement("cmd")
                    StepNode.appendChild(CmdNode)
                    CmdTextNode=CurrentOutputDOM.createTextNode("Place steps here")
                    CmdNode.appendChild(CmdTextNode)

                else: # concept is the default
                    # create the output body as conbody
                    CurrentOutputBody=CurrentOutputDOM.createElement("conbody")
                    TopicElement.appendChild(CurrentOutputBody)

            else:
                # process an element that is not temp:title
                # copy it into CurrentOutputBody
                CopyToOutput(node,CurrentOutputBody)
        else:
            debug(3,"Non-element node found in the main body, type="+str(node.nodeType))

# Postprocess - change links
# IMPORTANT: this must be run after all ID manupulation!

def PostprocessLinks(DOM):
    #print BookmarksDictionary
    # fix references for duplicate bookmarks
    for BookmarkID in BookmarksDuplicate:
        thedup=BookmarksDuplicate[BookmarkID]
        if thedup<>None:
            try:
                BookmarksDictionary[BookmarkID]=BookmarksDictionary[thedup]
            except KeyError:
                # something has failed, never mind, the link will break
                pass
    
    LinkNodes=DOM.getElementsByTagName("xref")
    NodesToDelete=set([])
    CurrentDocID=GetID(DOM)
    for LNode in LinkNodes:
        if LNode.getAttribute("scope")=="local":
            BookmarkID=LNode.getAttribute("href")
            try:
                (DocID,LinkID)=BookmarksDictionary[BookmarkID]
                if DocID==CurrentDocID:
                    LinkString=LinkID
                else:
                    LinkString=DocID+".dita"+LinkID
                LNode.setAttribute("href",LinkString)
            except KeyError:
                debug(1,"Bad bookmark "+BookmarkID)
                NodesToDelete.add(LNode)
                #print LNode.getAttribute("href")
    #print "DELETING NODES..."
    global Process_DeleteBadLinks
    if Process_DeleteBadLinks:
        for DNode in NodesToDelete:
            #print DNode.getAttribute("href")
            #print DNode.parentNode.tagName
            DestroyNode(DNode)
                

# Postprocess - remove otherprops attributes
def PostprocessRemoveOtherprops(DOM):

    # Internal recursion routine
    def RemoveOtherprops(node):
        if node.nodeType==xml.dom.Node.ELEMENT_NODE:
            if node.hasAttribute("otherprops"):
                node.removeAttribute("otherprops")
            for childnode in node.childNodes:
                RemoveOtherprops(childnode)
    for childnode in DOM.childNodes:
        RemoveOtherprops(childnode)


# Postprocess - attempt to create steps for a task. Run on split DOMs only
def PostprocessTaskSteps(DOM):
    taskElement=GetElementChild(DOM)
    if taskElement.tagName<>"task":
        return # not a task

    contextElement=taskElement.getElementsByTagName("context").item(0)
    stepsElement=taskElement.getElementsByTagName("steps").item(0)

    # newconElement and destroyElement are hacks.
    # newconelement will contain all nodes to remain in the context element
    # at the end the nodes are moved back into the context element and newconElement is destroyed
    # destroyelement iwll take all nodes to be destroyed, and it will just be destroyed with them
    # In this wayu we can iterate over the contents of content Element by referring to the first child every time
    # This hack is required cecause iterating over child nodes fails when some are moved in the process

    newconElement=DOM.createElement("newcon")
    taskElement.appendChild(newconElement)
    destroyElement=DOM.createElement("destroy")
    taskElement.appendChild(destroyElement)


    dummystepElement=GetElementChild(stepsElement)
    Started=False

#    for node in contextElement.childNodes:     # this does NOT work correctly

    while contextElement.hasChildNodes():
        node=contextElement.firstChild

        if (node.nodeType==xml.dom.Node.ELEMENT_NODE) and (node.tagName=="ol"):
            destroyElement.appendChild(node)
            for linode in node.childNodes: # these should be <li> and non-empty and anything else gets ignored
                if (linode.nodeType==xml.dom.Node.ELEMENT_NODE) and (linode.tagName=="li") and (linode.hasChildNodes()):

                    Started=True

                    # make the new step and its contents, delay info until needed
                    newstepElement=DOM.createElement("step")
                    stepsElement.insertBefore(newstepElement,dummystepElement)
                    cmdElement=DOM.createElement("cmd")
                    newstepElement.appendChild(cmdElement)
                    infoElement=None

                    # if the li node has an id, move it to the step node
                    if linode.hasAttribute("id"):
                        newstepElement.setAttribute("id",linode.getAttribute("id"))
                        linode.removeAttribute("id")

                    # if the first child node of the <li> node is a text node, move to cmd
                    if linode.firstChild.nodeType==xml.dom.Node.TEXT_NODE:
                        textN=linode.firstChild
                        if textN.data[-1].isspace():  # reomve last space - if no info this does not matter, and info ads a space
                            textN.data=textN.data[:-1]
                        cmdElement.appendChild(linode.firstChild)

                    # now the rest of the child nodes, if any, go into info
                    while linode.hasChildNodes():
                        if infoElement==None:
                            infoElement=DOM.createElement("info")
                            newstepElement.appendChild(infoElement)
                        infoElement.appendChild(linode.firstChild)
        elif Started: #everything, just everything, that is not an ordered list, goes into the last info
            if infoElement==None:
                infoElement=DOM.createElement("info")
                newstepElement.appendChild(infoElement) # there is a newstepElement as Started is true
            infoElement.appendChild(node)
        else:
            newconElement.appendChild(node)

    # now remove the dummy step if needed
    if Started:
        DestroyNode(dummystepElement)

    # Destroy the node with all stuff to be destroyed in it
    DestroyNode(destroyElement)

    # all elements to remain in context go back there, newcon gets destroyed
    while newconElement.hasChildNodes():
        contextElement.appendChild(newconElement.firstChild)
    DestroyNode(newconElement)
        

                        

                    
                
    

                
# ==========FINAL PROCESSING ======================
# This processing must be done AFTER postprocessing as it may make changes incompatible with postprocessing

# Rename all tags with a certain name to another name
def FinalTagRename(DOM,OldName,NewName):
    NodesToRename=[]
    for RNode in DOM.getElementsByTagName(OldName):
        NodesToRename.append(RNode)
    for RNode in NodesToRename:
        RNode.tagName=NewName

# Remove a tag, moving all text out of it
def FinalTagRemove(DOM,Name):
    TNodes=DOM.getElementsByTagName(Name)
    NodesToRemove=[]
    for node in TNodes:
        while node.hasChildNodes():
            node.parentNode.insertBefore(node.firstChild,node)
        NodesToRemove.append(node)
    for node in NodesToRemove:
        DestroyNode(node)
        
        

    
# === WRITING OUT ===============
# And creating the DITA Map
# note that we support a postfix on filenames - for multiple test writeouts mid-postprocess
def WriteOut(dir,postfix=''):
    # Create the DITAMAP DOM
    impl=xml.dom.getDOMImplementation("minidom")
    #doctype=impl.createDocumentType("map", "-//IBM//DTD DITA IBM Map//EN", "ibm-map.dtd")
    # OASIS
    doctype=impl.createDocumentType("map", "-//OASIS//DTD DITA Map//EN", "map.dtd")

    MapDOM=impl.createDocument(None,"map",doctype)

    # get the <map> element
    # FIXME: is there a better way to get the initially created element?
    MapNode=MapDOM.lastChild

    # Set <map> attributes
    MapNode.setAttribute("xml:lang","en-us")
    MapNode.setAttribute("title",DocumentTitle)
    #OASIS seems to require xmlns:ditaarch
    MapNode.setAttribute("xmlns:ditaarch","http://dita.oasis-open.org/architecture/2005/")


    # LevelDic contains last created topicref elements of all levels
    # For level 0 it has the <map> element
    LevelDic={0:MapNode}

    # LastLevel is the level for which the last topicref was created
    # A next level can not be over LastLevel+1 - we can not create empty topicrefs
    LastLevel=0

    # Skip first output as set
    global OutputDOMs
    OutputDOMs=OutputDOMs[SkipOutputSections:]

    # Write out topic DOMs while creating the Map DOM
    for DOM,level in OutputDOMs:

        # skip a DOM if it is textless
        if FindFirstText(DOM)==None:
            continue
        
        filename=GetID(DOM)+postfix+".dita"

        # Write the topic DOM
        outfile=codecs.open(os.path.join(dir,filename),"w",encoding="utf-8")
#        if UsePrettyXml:
#            xmlstr=DOM.toprettyxml(encoding="utf-8")
#        else:
        xmlstr=DOM.toxml("utf-8")
        outfile.write(xmlstr.decode("utf-8"))
        outfile.close()

        # Create topicref
        TopicrefNode=MapDOM.createElement("topicref")
        TopicrefNode.setAttribute("href",filename)
        TopicrefNode.setAttribute("format","dita")
        

        # Fix level if it is too high or too low
        if level>(LastLevel+1):
            level=LastLevel+1
        if level<1:
            level=1

        # Insert topicref under appropriate element
        LevelDic[level-1].appendChild(TopicrefNode)

        # Save the node in the dic, and save the last level
        LevelDic[level]=TopicrefNode
        LastLevel=level
        
    # Write out the ditamap

    mapfilename=os.path.join(dir,NameRoot+postfix+".ditamap")
    outfile=codecs.open(mapfilename,"w",encoding="utf-8")
#    if UsePrettyXml:
#        xmlstr=MapDOM.toprettyxml(encoding="utf-8")
#    else:
    xmlstr=MapDOM.toxml("utf-8")
#    xmlstr=xmlstr.decode("utf-8").encode("utf-8")
    outfile.write(xmlstr.decode("utf-8"))
    outfile.close()
    
    

# == MAIN CONVERSION PROCEDURE  =============================

def ConvertODTToDITA():
    try:
        # create the initial output DOM

        CreateInitialOutputDOM()

        # Open the zip file
        ODTzip=zipfile.ZipFile(InputFile)
        NamesInZip=ODTzip.namelist()


        # process the meta.xml file for document title
        MetaDOM=xml.dom.minidom.parse(ODTzip.open("meta.xml"))
        rootnode=MetaDOM.firstChild
        for childnode in rootnode.childNodes:
            if (childnode.nodeType==xml.dom.Node.ELEMENT_NODE):
                if childnode.tagName=="office:meta":
                    for grandchildnode in childnode.childNodes:
                        if (grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE):
                            if grandchildnode.tagName=="dc:title" and grandchildnode.hasChildNodes():
                                textnode=grandchildnode.firstChild
                                if textnode.nodeType==xml.dom.Node.TEXT_NODE:
                                    DocumentTitle=textnode.data


        # process the styles.xml file
        StylesDOM=xml.dom.minidom.parse(ODTzip.open("styles.xml"))
        rootnode=StylesDOM.firstChild
        for childnode in rootnode.childNodes:
            if (childnode.nodeType==xml.dom.Node.ELEMENT_NODE):
                if (childnode.tagName=="office:styles") or (childnode.tagName=="office:automatic-styles"):
                    ProcessStylesNode(childnode)

        # process the content.xml file
                
        ContentsDOM=xml.dom.minidom.parse(ODTzip.open("content.xml"))
        rootnode=ContentsDOM.firstChild
        for childnode in rootnode.childNodes:
            if (childnode.nodeType==xml.dom.Node.ELEMENT_NODE):
                if childnode.tagName=="office:automatic-styles":
                    ProcessStylesNode(childnode)
                elif childnode.tagName=="office:body":
                    #FIXME this is an outright hack - looking for text nodes only within office:body
                    # probably should be a separate function            
                    for grandchildnode in childnode.childNodes:
                        if (grandchildnode.nodeType==xml.dom.Node.ELEMENT_NODE):
                            if grandchildnode.tagName=="office:text":
                                ProcessTextNode(grandchildnode)



        # Remove trademark signs
        #ReplaceText(InitialOutputDOM,"®","")
        ReplaceText(InitialOutputDOM,u'®',"")

        # TEMP
        #outfile=open("test_prepost.dita","w")
        #xmlstr=InitialOutputDOM.toxml("UTF-8")
        #outfile.write(xmlstr)
        #outfile.close()


        PostprocessSpace(InitialOutputDOM)
        PostprocessLinebreak(InitialOutputDOM)


#        PostprocessElevateTempTopic(InitialOutputDOM)
        PostprocessFixNestedParagraphs(InitialOutputDOM)



        PostprocessFirstP(InitialOutputDOM)
        PostprocessFootnotes(InitialOutputDOM)
        PostprocessJoinLists(InitialOutputDOM)
        PostprocessNestedOneLists(InitialOutputDOM)

        PostprocessJoinTags(InitialOutputDOM,"i")
        PostprocessJoinTags(InitialOutputDOM,"u")
        PostprocessJoinTags(InitialOutputDOM,"codeph")

        # TEMP
        #outfile=open("dumpdump.dita","w")
        #xmlstr=InitialOutputDOM.toxml("UTF-8")
        #outfile.write(xmlstr)
        #outfile.close()


        PostprocessTables(InitialOutputDOM)
        PostprocessCodeblock(InitialOutputDOM)

        # TEMP
        #outfile=open("dumpdump.dita","w")
        #xmlstr=InitialOutputDOM.toxml("UTF-8")
        #outfile.write(xmlstr)
        #outfile.close()

        RemoveChildlessElementChildrenPara(InitialOutputDOM) # Remove blank paragraphs, now that they have been used for codeblocks

        # TEMP2
        #outfile=open("dumpdump2.dita","w")
        #xmlstr=InitialOutputDOM.toxml("UTF-8")
        #outfile.write(xmlstr)
        #outfile.close()


        PostprocessNotes(InitialOutputDOM)
        PostprocessNotesFromParagraphs(InitialOutputDOM)

        # TEMP2
        #outfile=open("dumpdump2.dita","w")
        #xmlstr=InitialOutputDOM.toxml("UTF-8")
        #outfile.write(xmlstr)
        #outfile.close()



        PostprocessBreakupIntoTopics()
        #debug(3,"Topic breakup complete")

        # TEMP
        #outfile=open("testnewdesign.dita","w")
        #xmlstr=InitialOutputDOM.toxml("UTF-8")
        #outfile.write(xmlstr)
        #outfile.close()

        #int("causeexception")


        # postprocess by DOM

        for DOM,level in OutputDOMs:
            PostprocessLinks(DOM)
            #print "Links has run"
            PostprocessRemoveOtherprops(DOM)
            #debug(3,"Otherprops has run")
        #debug(3,"Links/otherprops done")

            # FIXME: make this switchable
            #FinalTagRemove(DOM,"codeph")
            if TaskPost:
                PostprocessTaskSteps(DOM)


            # tag replacement
            if ReplaceBoldWith<>"":
                for DOM,level in OutputDOMs:
                    FinalTagRename(DOM,"b",ReplaceBoldWith)
            if ReplaceItalicWith<>"":
                for DOM,level in OutputDOMs:
                    FinalTagRename(DOM,"i",ReplaceItalicWith)
                
                               
        # write out
        try:
            WriteOut(OutputDirectory)
        except IOError as TheException:
            debug(0,"Error writing files!"+str(TheException))
            return

        # extract pictures, converting as necessary
        #print "Images started"
        # 0.35 conversion removed
        # 0.38 extracting formulas as in ToUnzip added
#        import Image
        for name in NamesInZip:
            if name.find("Pictures")==0:
                ODTzip.extract(name,OutputDirectory)
            else:
                for unzname in ToUnzip:
                    if name.find(unzname)>-1:
                        ODTzip.extract(name,OutputDirectory)
                        break
                    
#                if name[len(name)-4:]==".png":
#                    inname=os.path.join(OutputDirectory,name)
#                    outname=inname[:len(inname)-4]+".jpg"
#                    Image.open(inname).convert("RGB").save(outname,quality=95)

        debug(0,"Conversion completed successfully")
   #use this line to validate the try statement if exceptions are re-enabled temporarily
   #(by commenting out the "except Exception" clause)
#    except ValueError: pass

    except Exception as TheException:
        debug(0, "Exception "+str(type(TheException)))
        debug(0,str(TheException))

# ============ USER INTERFACE ================

# text viewer for log - copied from Python standard library

class TextViewer(Toplevel):
    """ Copied from: A simple text viewer dialog for IDLE
   
    
    """
    def __init__(self, parent, title, text):
        """Show the given text in a scrollable window with a 'close' button

        """
        Toplevel.__init__(self, parent)
        self.configure(borderwidth=5)
        self.geometry("=%dx%d+%d+%d" % (625, 500,
                                        parent.winfo_rootx() + 10,
                                        parent.winfo_rooty() + 10))
        #elguavas - config placeholders til config stuff completed
        self.bg = '#ffffff'
        self.fg = '#000000'

        self.CreateWidgets()
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.Ok)
        self.parent = parent
        self.textView.focus_set()
        #key bindings for this dialog
        self.bind('<Return>',self.Ok) #dismiss dialog
        self.bind('<Escape>',self.Ok) #dismiss dialog

        self.textView.config(font=tkFont.Font(size=16))

        self.textView.insert(0.0, text)
        self.textView.config(state=DISABLED)
        self.wait_window()

    def CreateWidgets(self):
        frameText = Frame(self, relief=SUNKEN, height=700)
        frameButtons = Frame(self)
        self.buttonOk = Button(frameButtons, text='Close',
                               command=self.Ok, takefocus=FALSE)
        self.scrollbarView = Scrollbar(frameText, orient=VERTICAL,
                                       takefocus=FALSE, highlightthickness=0)
        self.textView = Text(frameText, wrap=WORD, highlightthickness=0,
                             fg=self.fg, bg=self.bg)
        self.scrollbarView.config(command=self.textView.yview)
        self.textView.config(yscrollcommand=self.scrollbarView.set)
        self.buttonOk.pack()
        self.scrollbarView.pack(side=RIGHT,fill=Y)
        self.textView.pack(side=LEFT,expand=TRUE,fill=BOTH)
        frameButtons.pack(side=BOTTOM,fill=X)
        frameText.pack(side=TOP,expand=TRUE,fill=BOTH)

    def Ok(self, event=None):
        self.destroy()



# ============ MAIN CLASS
class Application(Frame):
    def Run(self):
        if InputFile=="NOT SELECTED":
            debug(0,"You must select input file")
            return

        if OutputDirectory=="NOT SELECTED":
            debug(0,"You must select output directory")
            return

        self.InFileButton.config(state="disabled")
        self.OutDirButton.config(state="disabled")
        self.RunButton.config(state="disabled")
        self.CloseButton.config(state="disabled")

        global Process_AntiquaAsBold
        Process_AntiquaAsBold=(self.UseAntiqueVar.get()==1)
        global Process_DeleteBadLinks
        Process_DeleteBadLinks=(self.DelbadlinksVar.get()==1)

        global DoNotPrefix
        DoNotPrefix=(self.DoNotPrefixVar.get()==1)

        global TaskPost
        TaskPost=(self.TaskPostVar.get()==1)

        global AggressiveFormula
        AggressiveFormula=(self.AggressiveFormulaVar.get()==1)



#        global UsePrettyXml
#        UsePrettyXml=(self.UsePrettyXmlVar.get()==1)

        global FrameMode
        FrameMode=(self.FrameModeVar.get()==1)

        global ReplaceBoldWith
        if self.ReplaceBoldVar.get()==1:
            ReplaceBoldWith=self.ReplaceBoldEntry.get()
        else:
            ReplaceBoldWith=""
            
        global ReplaceItalicWith
        if self.ReplaceItalicVar.get()==1:
            ReplaceItalicWith=self.ReplaceItalicEntry.get()
        else:
            ReplaceItalicWith=""      



        # reset globals

        global DebugText
        DebugText=""
        
        global BookmarksDictionary
        BookmarksDictionary={}

        global BookmarksDuplicate
        BookmarksDuplicate={}

        global StylesDictionary
        StylesDictionary={"":set([])}

#        global StylesOutlineLevels
#        StylesOutlineLevels={"":0}

        global InTable
        InTable=0


        global ListStylesDictionary
        ListStylesDictionary={"":{}}

        global CurrentListLevel
        CurrentListLevel=0

        global ListStyleNamesStack
        ListStyleNamesStack=[""]

        global OutputDOMs
        OutputDOMs=[]

        global AutoCreatedIDs
        AutoCreatedIDs=[]

        global AllIDs
        AllIDs=[]

        global AutoID_Counter
        AutoID_Counter=0

        global InitialOutputDOM
        InitialOutputDOM=None

        global InitialOutputBody
        InitialOutputBody=None

        global CurrentOutputTitle
        CurrentOutputTitle=None

        global CurrentOutputKeywords
        CurrentOutputKeywords=None

        global FrameModeStylesDictionary
        FrameModeStylesDictionary={}

        global ToUnzip
        ToUnzip=set([])




        ConvertODTToDITA()

        TextViewer(self,"Conversion log",DebugText)

        self.InFileButton.config(state="normal")
        self.OutDirButton.config(state="normal")
        self.RunButton.config(state="normal")
        self.CloseButton.config(state="normal")


        
    

    def SelectInputFile(self):
        ftypes = [('Document', '.odt'),
             ('All Files', '.*')]

        d = tkFileDialog.askopenfilename(filetypes=ftypes)

        if d<>"":
            global InputFile
            InputFile=d
            self.InFileLabel["text"]=d
            (pth,nme)=os.path.split(d)
            global NameRoot
            (NameRoot,ext)=os.path.splitext(nme)
            
        


    def SelectOutputDirectory(self):
        d = tkFileDialog.askdirectory(initialdir=".")
        if d<>"":
            global OutputDirectory
            OutputDirectory=d
            self.OutDirLabel["text"]=d
            
        
    def createWidgets(self):
        self.TopLabel = Label(self)
        self.TopLabel["text"] = "ODT to DITA converter version 0.41 OASIS 1"
        self.TopLabel.pack()

        self.TopLabel2 = Label(self)
        self.TopLabel2["text"] = "BETA VERSION: limited distribution"
        self.TopLabel2.pack()

        self.TopLabel3 = Label(self)
        self.TopLabel3["text"] = "(C) Copyright IBM 2010, 2016"
        self.TopLabel3.pack()

        
        self.TopLabel4 = Label(self)
        self.TopLabel4["text"] = "Code: Mikhail Ramendik ramendim@ie.ibm.com"
        self.TopLabel4.pack()

#        self.TopLabel5 = Label(self)
#        self.TopLabel5["text"] = "Tivoli ID Center of Excellence: Ireland"
#        self.TopLabel5.pack()

        
        self.InFileFrame = Frame(self)
        self.InFileFrame.pack()

        self.InFileButton = Button(self.InFileFrame)
        self.InFileButton.config(text="ODT file")
        self.InFileButton["command"] = self.SelectInputFile
        self.InFileButton.pack({"side":"left"})
        self.InFileLabel = Label(self.InFileFrame)
        self.InFileLabel["text"]="NOT SELECTED"
        self.InFileLabel.pack({"side":"left"})

        self.OutDirFrame = Frame(self)
        self.OutDirFrame.pack()

        self.OutDirButton = Button(self.OutDirFrame)
        self.OutDirButton.config(text="Output dir")
        self.OutDirButton["command"] = self.SelectOutputDirectory
        self.OutDirButton.pack({"side":"left"})
        self.OutDirLabel = Label(self.OutDirFrame)
        self.OutDirLabel["text"]="NOT SELECTED"
        self.OutDirLabel.pack({"side":"left"})

        self.UseAntiqueVar=IntVar()
        self.UseAntiqueCheck=Checkbutton(self)
        self.UseAntiqueCheck["text"]="Process Antiqua text as Bold"
        self.UseAntiqueCheck["variable"]=self.UseAntiqueVar
        self.UseAntiqueCheck.pack()

        self.DelbadlinksVar=IntVar()
        self.DelbadlinksCheck=Checkbutton(self)
        self.DelbadlinksCheck["text"]="Delete xref tags for non-working internal links"
        self.DelbadlinksCheck["variable"]=self.DelbadlinksVar
        self.DelbadlinksCheck.pack()

        self.FrameModeVar=IntVar()
        self.FrameModeCheck=Checkbutton(self)
        self.FrameModeCheck["text"]='Frame Mode: treat "headingN" style names as headers'
        self.FrameModeCheck["variable"]=self.FrameModeVar
        self.FrameModeCheck.pack()

        self.DoNotPrefixVar=IntVar()
        self.DoNotPrefixCheck=Checkbutton(self)
        self.DoNotPrefixCheck["text"]="Do NOT prefix file names with topic type letter"
        self.DoNotPrefixCheck["variable"]=self.DoNotPrefixVar
        self.DoNotPrefixCheck.pack()

        self.TaskPostVar=IntVar()
        self.TaskPostVar.set(1)
        self.TaskPostCheck=Checkbutton(self)
        self.TaskPostCheck["text"]="Process steps in tasks (uncheck if tasks are converted badly)"
        self.TaskPostCheck["variable"]=self.TaskPostVar
        self.TaskPostCheck.pack()

        self.AggressiveFormulaVar=IntVar()
        self.AggressiveFormulaCheck=Checkbutton(self)
        self.AggressiveFormulaCheck["text"]="Aggressive formula detection"
        self.AggressiveFormulaCheck["variable"]=self.AggressiveFormulaVar
        self.AggressiveFormulaCheck.pack()



#        self.UsePrettyXmlVar=IntVar()
#        self.UsePrettyXmlCheck=Checkbutton(self)
#        self.UsePrettyXmlCheck["text"]="Use pretty XML for output (UNTESTED)"
#        self.UsePrettyXmlCheck["variable"]=self.UsePrettyXmlVar
#        self.UsePrettyXmlCheck.pack()


        self.WarningLabel=Label(self)
        self.WarningLabel["text"]="Tag replacement: UNSUPPORTED. USE AT YOUR OWN RISK!"
        self.WarningLabel.pack()

        self.BoldFrame = Frame(self)
        self.BoldFrame.pack()
        self.ReplaceBoldVar=IntVar()
        self.ReplaceBoldCheck=Checkbutton(self.BoldFrame)
        self.ReplaceBoldCheck["text"]="Replace Bold tag with:"
        self.ReplaceBoldCheck["variable"]=self.ReplaceBoldVar
        self.ReplaceBoldCheck.pack({"side":"left"})
        self.ReplaceBoldEntry=Entry(self.BoldFrame)
        self.ReplaceBoldEntry.insert(0,"uicontrol")
        self.ReplaceBoldEntry.pack({"side":"left"})

        self.ItalicFrame = Frame(self)
        self.ItalicFrame.pack()
        self.ReplaceItalicVar=IntVar()
        self.ReplaceItalicCheck=Checkbutton(self.ItalicFrame)
        self.ReplaceItalicCheck["text"]="Replace Italic tag with:"
        self.ReplaceItalicCheck["variable"]=self.ReplaceItalicVar
        self.ReplaceItalicCheck.pack({"side":"left"})
        self.ReplaceItalicEntry=Entry(self.ItalicFrame)
        self.ReplaceItalicEntry.insert(0,"term")
        self.ReplaceItalicEntry.pack({"side":"left"})
        
        
        

        self.RunButton=Button(self)
        self.RunButton["text"]="Convert!"
        self.RunButton["command"]=self.Run
        self.RunButton.pack()

        self.CloseButton=Button(self, text="Close", command=root.destroy)
        self.CloseButton.pack()
        
        


        
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()


if len(sys.argv) == 1:
	# GUI
	root = Tk()
	app = Application(master=root)
	app.mainloop()
	try:
    		root.destroy()
	except Exception as e:
		print e 
		print "GUI fails, please use the command line. Use the -h option to view command line options"
else:
	#CLI
        parser = argparse.ArgumentParser("Convert ODT to DITA. (C) IBM, Mikhail Ramendik 2010, 2017")
        parser.add_argument("infile", help="the name/location of the input file")
        parser.add_argument("outdir", help="the name/location of the input file")
        parser.add_argument("-b", help="the tag to use for bold instead of b (optional)")
        parser.add_argument("-i", help="the tag to use for italic instead of i (optional)")
        parser.add_argument("-nt", help="do NOT try to process numbered staps in tasks (optional)", action="store_true")
        parser.add_argument("-np", help="do NOT prefix file names with topic type letter (optional)", action="store_true")
        parser.add_argument("-a", help="process antigua text as bold (optional)", action="store_true")
        parser.add_argument("-x", help="remove xref tags for non-working internal links (optional)", action="store_true")

        args = parser.parse_args()

        # assign input and output to globals
        InputFile=args.infile
        if not os.path.isfile(InputFile):
            print "Input file does not exist"
            sys.exit(2)

        (pth,nme)=os.path.split(InputFile)
        (NameRoot,ext)=os.path.splitext(nme)

        OutputDirectory=args.outdir
        if not os.path.isdir(OutputDirectory):
            print "Output directory does not exist, attempting to create"
            os.makedirs(OutputDirectory)
        


        # assign optional parameters to globals

        Process_AntiquaAsBold=args.a
        Process_DeleteBadLinks=args.x

        DoNotPrefix=args.np

        TaskPost=(not args.nt)

        if args.b != None:
            ReplaceBoldWith=args.b
        else:
            ReplaceBoldWith=""
            
        if args.i != None:
            ReplaceItalicWith=args.i
        else:
            ReplaceItalicWith=""

        print "Converting, please wait"

        ConvertODTToDITA()

        print "Conversion log:"
        print DebugText



	
    
