'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Functions for work with XML
'                                                 Kovalenko Dmitry. 16.08.2004

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
option explicit

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function xml_load_param_file(param_file_location,error_source)
 dim xml_doc
 set xml_doc=createobject("MSXML2.DOMDocument")

 call xml_doc.load(param_file_location)

 dim xml_err
 set xml_err=xml_doc.parseError

 if(xml_err.errorCode<>0)then
  dim msg

  msg="Error in parameter file ["&param_file_location&"]"&vbCrLf
  msg=msg&"reason:"&xml_err.reason&vbCrLf
  msg=msg&"line:"&cstr(xml_err.line)&" col:"&cstr(xml_err.linepos)&vbCrLf

  call err.raise(-1,error_source,msg)
 end if

 set xml_load_param_file=xml_doc
end function 'xml_load_param_file

'-------------------------------------------------------------------------
' Writed by Andrew Volkov for parsing an non-unicode files

function xml_load_param_file_not_unicode(param_file_location,error_source)
 dim xml_doc
 set xml_doc=createobject("MSXML2.DOMDocument")

 dim fso
 set fso = createobject("Scripting.FileSystemObject")

 dim s
 s = fso.OpenTextFile(param_file_location).readAll

 call xml_doc.loadXML(s)

 dim xml_err
 set xml_err=xml_doc.parseError

 if(xml_err.errorCode<>0)then
  dim msg

  msg="Error in parameter file ["&param_file_location&"]"&vbCrLf
  msg=msg&"reason:"&xml_err.reason&vbCrLf
  msg=msg&"line:"&cstr(xml_err.line)&" col:"&cstr(xml_err.linepos)&vbCrLf

  call err.raise(-1,error_source,msg)
 end if

 set xml_load_param_file_not_unicode=xml_doc
end function 'xml_load_param_file

'-------------------------------------------------------------------------
function xml_get_node(x,tag_name,required)
 const c_err_source="xml_get_node"

 dim t

 set t=x.selectSingleNode(tag_name)

 if((t is nothing) and required)then
  call err.raise(-1,c_err_source,"Tag ["&tag_name&"] not found")
 end if

 set xml_get_node=t
end function 'xml_get_node

'-------------------------------------------------------------------------
function xml_get_node_list(x,tag_name,required)
 const c_err_source="xml_get_node_list"

 dim t

 set t=x.getElementsByTagName(tag_name)

 if((t is nothing) and required)then
  call err.raise(-1,c_err_source,"Tag ["&tag_name&"] not found")
 end if

 set xml_get_node_list = t
end function'xml_get_node_list

'-------------------------------------------------------------------------
function xml_read_prop_str(x,tag_name,required,byref value)
 xml_read_prop_str=false

 dim prop

 set prop=xml_get_node(x,tag_name,required)

 if(not(prop is nothing))then
  value=trim(prop.text)

  xml_read_prop_str=true
 end if
end function 'xml_read_prop_str

'-------------------------------------------------------------------------
function xml_read_prop_str_req(x,tag_name)
 call xml_read_prop_str(x,tag_name,true,xml_read_prop_str_req)
end function 'xml_read_prop_str_req

'-------------------------------------------------------------------------
function xml_read_prop_str_list(x,tag_name,required)
 dim prop

 set prop = xml_get_node_list(x,tag_name,required)

 dim tag,list
 set list = createobject("LCPI.DBS.Utility.Vector")

 if(not(prop is nothing))then
  for tag = 0 to prop.length-1
   list.pushback(trim(prop(tag).text))
  next'tag
 end if

 set xml_read_prop_str_list=list
end function'xml_read_prop_str_list

'-------------------------------------------------------------------------
'def_value        - value for unknown tag
'def_empty_value  - value for empty tag

'return :tag value

function xml_read_prop_str_ex(x,tag_name,def_value,def_empty_value)
 dim value

 if(not xml_read_prop_str(x,tag_name,false,value))then
  xml_read_prop_str_ex=def_value
 elseif(value="")then
  xml_read_prop_str_ex=def_empty_value
 else
  xml_read_prop_str_ex=value
 end if
end function 'xml_read_prop_str_ex

'-------------------------------------------------------------------------
function xml_read_prop_lng(x,tag_name,required,byref value)
 const c_err_source="xml_read_prop_lng"

 xml_read_prop_lng=false

 dim prop

 set prop=xml_get_node(x,tag_name,required)

 if(prop is nothing)then exit function

 dim t
 t=trim(prop.text)

 if(not IsNumeric(t))then
  call err.raise(-1,c_err_source,"Parameter ["&tag_name&"] is not number")
 end if

 value=clng(t)

 xml_read_prop_lng=true
end function 'xml_read_prop_lng

'-------------------------------------------------------------------------
function xml_read_prop_lng_req(x,tag_name)
 call xml_read_prop_lng(x,tag_name,true,xml_read_prop_lng_req)
end function 'xml_read_prop_lng_req

'-------------------------------------------------------------------------
function xml_read_text_lng(x,required,byref value)
 const c_err_source="xml_read_text_lng"

 xml_read_text_lng=false

 if(x is nothing)then
  if(required)then
   call err.raise(-1,c_err_source,"Can't read required parameter")
  end if

  exit function
 end if

 dim t
 t=trim(x.text)

 if(not IsNumeric(t))then
  call err.raise(-1,c_err_source,"Parameter ["&x.nodeName&"] is not number")
 end if

 value=clng(t)

 xml_read_text_lng=true
end function 'xml_read_text_lng

'-------------------------------------------------------------------------
function xml_read_prop_dbl(x,tag_name,required,byref value)
 const c_err_source="xml_read_prop_dbl"

 xml_read_prop_dbl=false

 dim prop

 set prop=xml_get_node(x,tag_name,required)

 if(prop is nothing)then exit function

 dim t
 dim err_code

 t=trim(prop.text)

 on error resume next

 value=cdbl(t)

 err_code=err.number

 call err.clear()

 on error goto 0

 if(err_code<>0)then
  call err.raise(-1,c_err_source,"Parameter ["&tag_name&"] is not real number!")
 end if

 xml_read_prop_dbl=true
end function 'xml_read_prop_dbl

'-------------------------------------------------------------------------
public function xml_read_prop_dbl_req(x,tag_name)
 call xml_read_prop_dbl(x,tag_name,true,xml_read_prop_dbl_req)
end function 'xml_read_prop_dbl_req

'-------------------------------------------------------------------------
public function xml_read_prop_date(x,tag_name,required,byref value)
 const c_err_source="xml_read_prop_date"

 xml_read_prop_date=false

 dim prop

 set prop=xml_get_node(x,tag_name,required)

 if(prop is nothing)then exit function

 dim t
 dim err_code

 t=trim(prop.text)

 on error resume next

 value=cdate(t)

 err_code=err.number

 call err.clear()

 on error goto 0

 if(err_code<>0)then
  call err.raise(-1,c_err_source,"Parameter ["&tag_name&"] is not date!")
 end if

 xml_read_prop_date=true
end function 'xml_read_prop_date

'-------------------------------------------------------------------------
function xml_read_prop_date_req(x,tag_name)
 call xml_read_prop_date(x,tag_name,true,xml_read_prop_date_req)
end function 'xml_read_prop_date_req

'*******************************************************************************

Function xml_read_attr_bool(x, attr_name, required, ByRef value)
  Const c_err_source = "xml_read_attr_bool"
  xml_read_attr_bool = False

  Dim attr_value
  If (Not xml_get_text_attr(x, attr_name, required, attr_value)) Then Exit Function

  Select Case attr_value
    Case "true"
      value = True
    Case "false"
      value = False
    Case Else
      Dim msg : msg = "Wrong value of attribute [" & attr_name & "]. Expected 'true' or 'false'!"
      Call err.raise(-1, c_err_source, msg)
  End Select

  xml_read_attr_bool = True
End Function  ' xml_read_attr_bool

'*******************************************************************************

Function xml_read_attr_bool_ex(x, attr_name, def_value)
  Dim value

  If (Not xml_read_attr_bool(x, attr_name, false, value)) Then
    xml_read_attr_bool_ex = def_value
  Else
    xml_read_attr_bool_ex = value
  End If
End Function  ' xml_read_attr_bool_ex

'*******************************************************************************

Function xml_read_attr_lng(x,attr_name,required,byref value)
  Const c_err_source = "xml_read_attr_lng"
  xml_read_attr_lng = False

  Dim attr_value
  If (Not xml_get_text_attr(x, attr_name, required, attr_value)) Then Exit Function

  If (Not IsNumeric(attr_value)) Then
    Call err.raise(-1,c_err_source, "Attribute [" & attr_name & "] is not number!")
  End If

  value = CLng(attr_value)
  xml_read_attr_lng = True
End Function  ' xml_read_attr_lng

'*******************************************************************************

Function xml_read_attr_str(x, attr_name, required, ByRef value)
  Const c_err_source = "xml_read_attr_str"

  xml_read_attr_str = False

  Dim attr_value
  If (Not xml_get_text_attr(x, attr_name, required, attr_value)) Then Exit Function

  value = attr_value
  xml_read_attr_str = True
End Function 'xml_read_attr_str

'*******************************************************************************

Function xml_read_attr_str_ex(x,attr_name,def_value,def_empty_value)
 dim value

 if(not xml_read_attr_str(x,attr_name,false,value))then
  xml_read_attr_str_ex=def_value
 elseif(value="")then
  xml_read_attr_str_ex=def_empty_value
 else
  xml_read_attr_str_ex=value
 end if
end function 'xml_read_attr_str_ex

'*******************************************************************************
' Helper functions

' Getting an attribute value by name
Private Function xml_get_text_attr(x, attr_name, required, ByRef attr_value)
  Const c_err_source = "xml_get_text_attr"
  xml_get_text_attr = False

  Dim attr : set attr = x.attributes.getNamedItem(attr_name)

  If (attr Is Nothing) Then
    If (required) Then
      Call err.raise(-1, c_err_source, "Attribute [" & attr_name & "] not found!")
    End If

    Exit Function
  End If

  attr_value = attr.text
  xml_get_text_attr = True
End Function  ' xml_get_text_attr

'*******************************************************************************
