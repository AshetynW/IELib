/*
This AHK Library is intended for use when utilizing COM Methods to Internet Explorer Applications. For support please reach out to me at Ashetynw@Gmail.com

*IMPORTANT*
Before utilizing this library, the IE object must be invoked via the "IEGet" Function. Please ensure that you are making your IE object handle a global var prior to use in this library.

Version: 2
*/

IEGet(Name=""){       ;Retrieve pointer to existing IE window/tab. Pass the title of the IE page to this function to initialize it as an object. example, ie:=IEGet("Google")

    IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
        Name := ( Name="New Tab - Windows Internet Explorer" ) ? "about:Tabs"
        : RegExReplace( Name, " - (Windows|Microsoft) Internet Explorer" )
    For wb in ComObjCreate( "Shell.Application" ).Windows
        If ( wb.LocationName = Name ) && InStr( wb.FullName, "iexplore.exe" )
            Return wb
}

FindIE(ShowMessage, Title){ ;Pass "True/False" to the Showmessage arg and the title of your page to the Title Arg. This will attempt to make the page the IE object, if not able, the item not found tip will show.
 ie := IeGet(Title)
 if (IsObject(ie)=false) &&(ShowMessage=True){
 TrayTip,Item List not found, Can't find the Item List screen!
    return false
  } 
 return ie
}

IeLoad(Browser, Url:=""){
  ;Function Details:
  ;Returns True or False. Will indicate if the process was successful
  ;Parameter 1: Send a pointer to an internet explorer browser
  ;See IEGet on how to get a pointer
  ;Function will wait for a page to load
  If !Browser
   return false
  
  ;Wait up to 2 minutes
  if (URL!="")
  {
   Loop, 480
   {
    sleep 250
    if (Instr(Browser.LocationURL,Url))
     break   
   }
  }
  
  loop, 50
  {
   if (Browser.busy || Browser.ReadyState<>4)
    sleep 100
  }
  
  loop, 4
  {
   tagnames:=Browser.document.getelementsbytagname("*").Length
   loop, 10
   {
    Sleep 100
    if (tagnames = Browser.document.getelementsbytagname("*").Length)
     return true
   }
  }
  
  return false
 }
 
ElementInteraction(ElementID, ElementTag, Browser, Action, ValueToSet, Type){ 
	/*
	Arguments are utilized by DOM Elements in IE
	ElementID: this is the specific Identifier handle you're looking for. This can either be the "Name" or the "ID" of the element itself. Case Sensitive
	ElementTag: This is the actual tag that houses the Element. So, <div>; <input>; <p> ; <a> etc...
	Browser: This is the IE Object Handle. Typically made as "ie"
	Action: The desired action once the object has been found. Accepted inputs are "Click", "Set", and "Get"
	ValueToSet: This is only utilized in the "Set" Action. This must be not null to pass appropriately, but will not affect the "Get" and "Click" Portions
	Type: Specifies the type of input you are using. Currently, this function can only be utilized with "ID" or "Name"
	*/
   
   IfWinNotExist, File Explorer
		run, explorer.exe
		
   if (Type = "ID") 
    {
		try
		{
			AllItems:=ie.Document.getElementsByTagName("*")
			DomObj:=ie.Document
				
			Length:=AllItems.Length
				
			loop, % Length
			{
					ToolTip, %A_index%
				if (instr(AllItems[A_Index-1].ID,ElementID)){
					my_element:=DomObj.getElementById(ElementID)
					
					if (Action="Click"){
						WinActivate, File Explorer
						WinMinimize, File Explorer
						clickLast(ElementID, ie, false)
						my_element.focus
						my_element.click()
						my_element.onClick()
						my_element.FireEvent("OnClick")
						Sleep 600
						IeLoad(Browser)
						return True
					}
						   
					if (Action="Set"){
						my_element.focus
						my_element.value:=ValueToSet
						Sleep, 50
						my_element.onChange()
						my_element.FireEvent("onChange")
						IeLoad(Browser)
						return True
					}
						   
					if (Action="Get"){
						return my_element.Value
					}
				}
			}
		}
	}
	
	if (Type = "Name") {
		Try
		{
			AllItems:=ie.Document.GetElementsByTagName(ElementTag)
			DomObj:=ie.Document
			
			Length:=AllItems.Length
			
			Loop % Length
			{
				Haystack:= AllItems[A_index-1].OuterHtml
				Needle =  name=`"%ElementID%`"
				
				if (instr(AllItems[A_index-1].OuterHtml,ElementID))
				{
					my_element:=DomObj.getElementsByTagName(ElementTag)[A_index-1]
					ClickElement:=DomObj.getElementsByTagName(ElementID)
					
					if (Action="Click"){
						WinActivate, File Explorer
						WinMinimize, File Explorer
						clickLast(ElementID, ie, false)
						my_element.focus
						my_element.click()
						my_element.onClick()
						my_element.FireEvent("OnClick")
						Sleep 600
						IeLoad(Browser)
						return True
					}
						   
					if (Action="Set"){
						my_element.focus
						my_element.value:=ValueToSet
						Sleep, 50
						my_element.onChange()
						my_element.FireEvent("onChange")
						IeLoad(Browser)
						return True
					}
						   
					if (Action="Get"){
						return my_element.Value
					}
					
				}
			}
		}
	}
 return false
}

ClickLast(ElementID, Browser, iframe){ ;Still Under Construction
	
	if (iframe = TRUE) 
	{
		AllItems:=Browser.Document.getElementsByTagName("iFrame")
		Length:=AllItems.Length
	 
		loop, % Length
		{
			if (instr(AllItems[A_Index-1].ID,ElementID) && isObject(AllItems[A_Index-1].contentdocument.getElementById(ElementID))){
				AllItems[A_Index-1].contentdocument.getElementById(elementID).Click()
				IeLoad(Browser)
				return True
			} 
		}
	} else {
		AllItems:=Browser.Document.getElementsByTagName("*")
		Length:=AllItems.Length
		DomObj:=ie.Document
		
		loop, % Length
		{
			if (instr(AllItems[A_Index-1].ID,ElementID) && isObject(DomObj.getElementById(ElementID))){
				DomObj.getElementById(elementID).Click()
				IeLoad(Browser)
				return True
			} 
		}
	}
 return False
}

FindElement(Document, ElementName, ElementType, Index:=0, TagName:="*", LikeMatch:=False){ ; Still Under Construction
 /*
  Function Details:
   Returns an Element Object
   Parameter 1: Pointer to an IE.Document
   Parameter 2: The ElementName you are searching for
    If you are searching for the: text, alt, value, innertext or title properties this is value you are looking for
   Parameter 3: The ElementType you are looking for
    Valid Parameters: ID, Name, ClassName,Text, Alt, Value, InnerText or Title. All parameters are strings.
   Parameter 4 (Optional): This is the index number for getElementsByName/ClassName. 
    Default value is 0 (first item)
   Parameter 5 (Optional): This is the TagName keyword to reduce the search space in getElementsbyTagName. 
    Default value is everything ("*")
 */

  loop, 10
  {
   try
   {
    SubDocument:= Document.getElementsbyTagName("iFrame")[A_Index-1].contentdocument
    FindElement:=FindElement(SubDocument,ElementName,ElementType,Index,TagName,LikeMatch)
    if isObject(FindElement){
     FindElement.Focus()
     return FindElement
    }
   }
  } 

  if (LikeMatch=True){
   AllItems:=Document.Body.getElementsbyTagName(TagName)
   Length:=AllItems.Length

   Loop, %Length%
   {
    try
    {
     if (ElementType="ID" && instr(AllItems[A_Index-1].ID,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="Name" && instr(AllItems[A_Index-1].Name,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="ClassName" && instr(AllItems[A_Index-1].Class,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="Value" && instr(AllItems[A_Index-1].Value,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="Alt" && instr(AllItems[A_Index-1].Alt,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="Text" && instr(AllItems[A_Index-1].Text,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="Title" && instr(AllItems[A_Index-1].Title,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="InnerText" && instr(AllItems[A_Index-1].InnerText,ElementName))
      return AllItems[A_Index-1]
     
     if (ElementType="Href" && instr(AllItems[A_Index-1].href,ElementName))
      return AllItems[A_Index-1]
    }
   }
  } 
  
  if (LikeMatch=False){
    if (ElementType="Name" && isobject(Document.getElementsbyName(ElementName)[Index]))
     return Document.getelementsbyName(ElementName)[Index]
    if (ElementType="ClassName" && isObject(Document.getElementsbyClassName(ElementName)[Index]))
     return Document.getElementsbyClassName(ElementName)[Index]
    if (ElementType="ID" && isObject(Document.getElementbyID(ElementName)))
     return Document.getElementbyID(ElementName)
   
   AllItems:=Document.Body.getElementsbyTagName(TagName)
   Length:=AllItems.Length
   
   Loop, %Length%
   {
    try
    {
     if (ElementType="Text" && AllItems[A_Index-1].text=ElementName)
      return AllItems[A_Index-1]
     
     if (ElementType="Value" && AllItems[A_Index-1].value=ElementName)
      return AllItems[A_Index-1]
     
     if (ElementType="Alt" && AllItems[A_Index-1].alt=ElementName)
      return AllItems[A_Index-1]
     
     if (ElementType="Title" && AllItems[A_Index-1].title=ElementName)
      return AllItems[A_Index-1]
     
     if (ElementType="InnerText" && AllItems[A_Index-1].InnerText=ElementName)
      return AllItems[A_Index-1]
     
     if (ElementType="Href" && AllItems[A_Index-1].href=ElementName)
      return AllItems[A_Index-1]
     
     if (ElementType="TabIndex" && AllItems[A_Index-1].TabIndex=ElementName)
      return AllItems[A_Index-1]
    }
   }
  }
 }