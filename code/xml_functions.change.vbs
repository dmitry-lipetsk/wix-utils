'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function xml_add_element(doc,parent_node, name)
 dim x
 
 set x=doc.createElement(name)
 
 parent_node.appendChild(x)
 
 set xml_add_element=x
end function 'xml_add_element

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
function xml_add_element_with_text(doc,parent_node,name,text)
 dim x
 set x=xml_add_element(doc,parent_node,name)

 call x.appendChild(doc.createTextNode(text))
 
 set xml_add_element_with_text=x
end function 'xml_add_element_with_text
