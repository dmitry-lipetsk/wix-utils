option explicit

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'scan files database
'<scan_files>
' <header>
'  <source>..\..\lib</source>
'
'  OPTIONAL Filters for root directory
'  <include>name</include>
'
'  <gen_dir_id>0</gen_dir_id>
'  <gen_file_id>0</gen_file_id>
'  <target>x86</target>
'  <target>x64</target>
'  <target>x86_free<target>
' </header>
' <data>
'  <dir parent_id='parent_dir_number' id='unique_dir_number' name='name' exists='0|1'/>
'  <file parent_id='parent_dir_number' id='unique_file_number' name='name' exists='0|1'>
'   <guid target='target_name'>guid_string</guid>
'  </file>
' </data>
'</scan_files>

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
const c_xml_node__install_files  ="install_files"
const c_xml_node__header         ="header"
const c_xml_node__source         ="source"
const c_xml_node__include        ="include"
const c_xml_node__target         ="target"
const c_xml_node__gen_dir_id     ="gen_dir_id"
const c_xml_node__gen_file_id    ="gen_file_id"
const c_xml_node__data           ="data"
const c_xml_node__dir            ="dir"
const c_xml_node__file           ="file"
const c_xml_node__guid           ="guid"

const c_xml_attr__id             ="id"
const c_xml_attr__parent_id      ="parent_id"
const c_xml_attr__name           ="name"
const c_xml_attr__exists         ="exists"
const c_xml_attr__target         ="target"

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'class t_install_files_database__generator

class t_install_files_database__generator
 private m_value

 private sub class_initialize()
  call init(0)
 end sub 'class_initialize

 public sub init(value)
  m_value=value
 end sub 'init

 public property get current_value()
  current_value=m_value
 end property 'get current_value

 public function gen_id()
  m_value=m_value+1

  gen_id=m_value
 end function 'gen_id
end class 't_install_files_database__generator

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'class t_install_files_database__file

class t_install_files_database__file
 private m_parent_dir_id
 private m_id
 private m_name
 private m_exists
 private m_guids

 private sub class_initialize()
  set m_guids=nothing
 end sub

 public sub init(parent_dir_id,id,name)
  m_parent_dir_id=parent_dir_id
  m_id=id
  m_name=name
  m_exists=1
  set m_guids=createobject("Scripting.Dictionary")
 end sub 'init

 '------------------------------------------------------------------------
 public property get parent_dir_id()
  parent_dir_id=m_parent_dir_id
 end property

 '------------------------------------------------------------------------
 public property get id()
  id=m_id
 end property

 '------------------------------------------------------------------------
 public property get name()
  name=m_name
 end property

 '------------------------------------------------------------------------
 public property get exists()
  exists=m_exists<>0
 end property

 '------------------------------------------------------------------------
 public sub load_xml_data(xml_entry)
  const c_err_src="t_install_files_database__file::load_xml_data"

  call xml_read_attr_lng(xml_entry,c_xml_attr__parent_id,true,m_parent_dir_id)

  call xml_read_attr_lng(xml_entry,c_xml_attr__id,true,m_id)

  call xml_read_attr_str(xml_entry,c_xml_attr__name,true,m_name)

  call xml_read_attr_lng(xml_entry,c_xml_attr__exists,true,m_exists)

  set m_guids=createobject("Scripting.Dictionary")

  dim xguids
  set xguids=xml_get_node_list(xml_entry,c_xml_node__guid,false)

  if(not (xguids is nothing))then
   dim i,xnode__guid,target

   for i=0 to xguids.length-1
    set xnode__guid=xguids(i)

    call xml_read_attr_str(xnode__guid,c_xml_attr__target,true,target)

    if(m_guids.exists(target))then
     call err.raise(-1,c_err_src,"multiple guids for target ["&target&"]")
    end if

    call m_guids.add(target,xnode__guid.Text)
   next 'i
  end if
 end sub 'load_xml_data

 '------------------------------------------------------------------------
 public function create_xml_entry(xdoc)
  dim xnode__file
  set xnode__file=xdoc.createElement(c_xml_node__file)

  call xnode__file.setAttribute(c_xml_attr__parent_id,m_parent_dir_id)

  call xnode__file.setAttribute(c_xml_attr__id,m_id)

  call xnode__file.setAttribute(c_xml_attr__name,m_name)

  call xnode__file.setAttribute(c_xml_attr__exists,m_exists)

  dim target

  for each target in m_guids.keys
   call xnode__file.appendChild(create_xml_entry__guid(xdoc,target,m_guids.item(target)))
  next 'target

  set create_xml_entry=xnode__file
 end function 'create_xml_entry

 '------------------------------------------------------------------------
 private function create_xml_entry__guid(xdoc,target,guid)
  dim xnode__guid
  set xnode__guid=xdoc.createElement(c_xml_node__guid)

  call xnode__guid.setAttribute(c_xml_attr__target,target)
  call xnode__guid.appendChild(xdoc.createTextNode(guid))

  set create_xml_entry__guid=xnode__guid
 end function 'create_xml_entry__guid

 '------------------------------------------------------------------------
 public sub set_exists()
  m_exists=1
 end sub 'set_exists

 '------------------------------------------------------------------------
 public sub reset_exists()
  m_exists=0
 end sub 'reset_exists

 '------------------------------------------------------------------------
 public function reg_guid(target)
  if(m_guids.Exists(target))then
   reg_guid=m_guids.item(target)
  else
   dim guid

   guid=(createobject("Scriptlet.TypeLib")).Guid

   guid=mid(guid,2,36)

   guid=ucase(guid)

   call m_guids.add(target,guid)

   reg_guid=guid
  end if
 end function 'reg_guid

 '------------------------------------------------------------------------
 public function get_guid(target)
  if(not m_guids.Exists(target))then
   dim msg
   
   msg="file ["&m_id&"] not contains a guid for target ["&target&"]"
   
   call err.raise(-1,"t_install_files_database__file::get_guid",msg)
  end if 

  get_guid=m_guids.item(target)  
 end function 'get_guid
end class 't_install_files_database__file

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'class t_install_files_database__dir

class t_install_files_database__dir
 private m_parent_dir_id
 private m_id
 private m_name
 private m_exists

 public m_dirs
 public m_files

 private sub class_initialize()
  set m_dirs=nothing
  set m_files=nothing
 end sub

 public sub init(parent_dir_id,id,name)
  m_parent_dir_id=parent_dir_id
  m_id=id
  m_name=name
  m_exists=1
  set m_dirs=CreateObject("Scripting.Dictionary")
  set m_files=CreateObject("Scripting.Dictionary")
 end sub 'init

 '------------------------------------------------------------------------
 public property get parent_dir_id()
  parent_dir_id=m_parent_dir_id
 end property

 '------------------------------------------------------------------------
 public property get id()
  id=m_id
 end property

 '------------------------------------------------------------------------
 public property get name()
  name=m_name
 end property 

 '------------------------------------------------------------------------
 public property get exists()
  exists=m_exists<>0
 end property

 '------------------------------------------------------------------------
 public sub load_xml_data(xml_entry)
  if(not xml_read_attr_lng(xml_entry,c_xml_attr__parent_id,false,m_parent_dir_id))then
   m_parent_Dir_id=null
  end if

  call xml_read_attr_lng(xml_entry,c_xml_attr__id,true,m_id)

  call xml_read_attr_str(xml_entry,c_xml_attr__name,true,m_name)

  call xml_read_attr_lng(xml_entry,c_xml_attr__exists,true,m_exists)

  set m_dirs=CreateObject("Scripting.Dictionary")
  set m_files=CreateObject("Scripting.Dictionary")
 end sub 'load_xml_data

 '------------------------------------------------------------------------
 public function create_xml_entry(xdoc)
  dim xnode__dir
  set xnode__dir=xdoc.createElement(c_xml_node__dir)

  if(not IsNull(m_parent_dir_id))then
   call xnode__dir.setAttribute(c_xml_attr__parent_id,m_parent_dir_id)
  end if

  call xnode__dir.setAttribute(c_xml_attr__id,m_id)

  call xnode__dir.setAttribute(c_xml_attr__name,m_name)

  call xnode__dir.setAttribute(c_xml_attr__exists,m_exists)

  set create_xml_entry=xnode__dir
 end function 'create_xml_entry

 '------------------------------------------------------------------------
 public sub set_exists()
  m_exists=1
 end sub 'set_exists

 '------------------------------------------------------------------------
 public sub reset_exists()
  m_exists=0
 end sub 'reset_exists
end class 't_install_files_database__dir

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'class t_install_files_database

class t_install_files_database
 private m_source__path

 private m_includes
 
 private m_gen_dir_id
 private m_gen_file_id

 private m_targets

 private m_root_dir

 private m_all_dirs__by_id
 private m_all_files__by_id

 '------------------------------------------------------------------------
 private sub class_initialize()
  set m_includes         =nothing
  set m_gen_dir_id       =nothing
  set m_gen_file_id      =nothing
  set m_targets          =nothing
  set m_root_dir         =nothing
  set m_all_dirs__by_id  =nothing
  set m_all_files__by_id =nothing
 end sub 'class_initialize

 '------------------------------------------------------------------------
 public sub init(source__path)
  m_source__path=source__path

  set m_includes         =createobject("Scripting.Dictionary")

  set m_gen_dir_id  =new t_install_files_database__generator
  set m_gen_file_id =new t_install_files_database__generator

  set m_targets          =createobject("Scripting.Dictionary")
  set m_all_dirs__by_id  =createobject("Scripting.Dictionary")
  set m_all_files__by_id =createobject("Scripting.Dictionary")

  set m_root_dir=nothing
 end sub 'init

 '------------------------------------------------------------------------
 public property get source__path()
  source__path=m_source__path
 end property 'get source_path

 '------------------------------------------------------------------------
 public property get targets()
  targets=m_targets.keys
 end property 'get targets

 '------------------------------------------------------------------------
 public property get root_dir()
  if(m_root_dir is nothing)then
   dim tmp
   set tmp=new t_install_files_database__dir

   call tmp.init(null,m_gen_dir_id.gen_id(),"")

   call m_all_dirs__by_id.add(tmp.id,tmp)

   set m_root_dir=tmp
  end if

  call m_root_dir.set_exists()
  
  set root_dir=m_root_dir
 end property 'get root_dir

 '------------------------------------------------------------------------
 public property get all_files()
  all_files=m_all_files__by_id.items
 end property 'get files

 '------------------------------------------------------------------------
 public function get_root_dir()
  set get_root_dir=m_root_dir
 end function 'get_root_dir

 '------------------------------------------------------------------------
 public sub load_from_file(file_path)
  const c_err_src="t_install_files_database::load_from_file"

  dim xdoc
  set xdoc=xml_load_param_file(file_path,c_err_src)

  dim xnode__install_files
  set xnode__install_files=xml_get_node(xdoc,c_xml_node__install_files,true)

  dim xnode__header
  set xnode__header=xml_get_node(xnode__install_files,c_xml_node__header,true)

  call init(xml_read_prop_str_req(xnode__header,c_xml_node__source))

  'загрузка списка includes
  dim xnodes
  dim x

  set xnodes = xml_get_node_list(xnode__header,c_xml_node__include,true)

  if(not (xnodes is nothing))then
   for x=0 to xnodes.length-1
    call add_include(trim(xnodes(x).text))
   next 'x
  end if

  '---------
  call m_gen_dir_id.init(xml_read_prop_lng_req(xnode__header,c_xml_node__gen_dir_id))
  call m_gen_file_id.init(xml_read_prop_lng_req(xnode__header,c_xml_node__gen_file_id))

  'загрузка списка target-ов
  set xnodes = xml_get_node_list(xnode__header,c_xml_node__target,true)

  for x=0 to xnodes.length-1
   call add_target(trim(xnodes(x).text))
  next 'x

  dim xnode__data
  set xnode__data=xml_get_node(xnode__install_files,c_xml_node__data,true)

  'загрузка каталогов
  dim msg
  set xnodes=xml_get_node_list(xnode__data,c_xml_node__dir,false)

  dim dir

  if(not (xnodes is nothing))then
   for x=0 to xnodes.length-1
    set dir=new t_install_files_database__dir

    call dir.load_xml_data(xnodes(x))

    if(m_all_dirs__by_id.exists(dir.id))then
     msg="Multiple dirs with one id ["&dir.id&"]"
     call err.raise(-1,c_err_src,msg)
    end if

    call m_all_dirs__by_id.add(dir.id,dir)
   next 'x
  end if

  'загрузка файлов
  set xnodes=xml_get_node_list(xnode__data,c_xml_node__file,false)

  dim file

  if(not (xnodes is nothing))then
   for x=0 to xnodes.length-1
    set file=new t_install_files_database__file

    call file.load_xml_data(xnodes(x))

    if(m_all_files__by_id.exists(file.id))then
     msg="Multiple files with one id ["&file.id&"]"
     call err.raise(-1,c_err_src,msg)
    end if

    call m_all_files__by_id.add(file.id,file)
   next 'x
  end if

  'связывание каталогов
  dim parent_dir,dir2

  for each dir in m_all_dirs__by_id.items
   if(IsNull(dir.parent_dir_id))then
    if(not (m_root_dir is nothing))then
     msg="Multiple root dirs: ["&m_root_dir.id&"], ["&dir.id&"]"
     call err.raise(-1,c_err_src,msg)
    end if

    set m_root_dir=dir
   else
    if(not m_all_dirs__by_id.Exists(dir.parent_dir_id))then
     msg="Unknown parent_dir_id ["&dir.parent_dir_id&"] of dir [id: "&dir.id&"]"
     call err.raise(-1,c_err_src,msg)
    end if

    set parent_dir=m_all_dirs__by_id.Item(dir.parent_dir_id)

    if(parent_dir.m_dirs.Exists(dir.name))then
     msg="parent directory ["&parent_dir.id&"] already contains the dir with name ["&dir.name&"]. can't append the dir with id ["&dir.id&"]"
     call err.raise(-1,c_err_src,msg)
    end if

    if((dir.exists) and (not parent_dir.exists))then
     msg="parent dir ["&parent_dir.id&"] mark as not exists, when the dir ["&dir.id&"] mark as exists!"
     call err.raise(-1,c_err_src,msg)
    end if

    'detect "cycle" errors
    set dir2=parent_dir
    
    while(not IsNull(dir2.parent_dir_id))
     if(dir2.id=dir.id)then
      msg="detect the cycle with dir ["&dir.id&"]"
      call err.raise(-1,c_err_src,msg)
     end if
     
     if(not m_all_dirs__by_id.Exists(dir2.parent_dir_id))then
      msg="Unknown parent_dir_id ["&dir2.parent_dir_id&"] of dir [id: "&dir2.id&"]"
      call err.raise(-1,c_err_src,msg)
     end if

     set dir2=m_all_dirs__by_id.Item(dir2.parent_dir_id)
    wend
    
    call parent_dir.m_dirs.add(dir.name,dir)
   end if
  next 'dir

  'связывание файлов
  for each file in m_all_files__by_id.items
   if(not m_all_dirs__by_id.Exists(file.parent_dir_id))then
    msg="Unknown parent_dir_id ["&file.parent_dir_id&"] of file [id: "&file.id&"]"
    call err.raise(-1,c_err_src,msg)
   end if

   set parent_dir=m_all_dirs__by_id.Item(file.parent_dir_id)

   if(parent_dir.m_files.Exists(file.name))then
    msg="parent directory ["&parent_dir.id&"] already contains the file with name ["&file.name&"]. can't append the file with id ["&file.id&"]"
    call err.raise(-1,c_err_src,msg)
   end if

   if((file.exists) and (not parent_dir.exists))then
    msg="parent dir ["&parent_id.id&"] mark as not exists, when the file ["&file.id&"] mark as exists!"
    call err.raise(-1,c_err_src,msg)
   end if

    call parent_dir.m_files.add(file.name,file)
  next 'file
 end sub 'load_from_file

 '------------------------------------------------------------------------
 public sub save_to_file(file_path)
  dim xdoc
  set xdoc=createobject("MSXML.DOMDocument")

  dim xnode__install_files
  set xnode__install_files=xml_add_element(xdoc,xdoc,c_xml_node__install_files)

  dim xnode__header
  set xnode__header=xml_add_element(xdoc,xnode__install_files,c_xml_node__header)

  call xml_add_element_with_text(xdoc,xnode__header,c_xml_node__source,m_source__path)

  dim x
  for each x in m_includes.keys
   call xml_add_element_with_text(xdoc,xnode__header,c_xml_node__include,x)
  next

  call xml_add_element_with_text(xdoc,xnode__header,c_xml_node__gen_dir_id,m_gen_dir_id.current_value)
  call xml_add_element_with_text(xdoc,xnode__header,c_xml_node__gen_file_id,m_gen_file_id.current_value)

  for each x in m_targets.keys
   call xml_add_element_with_text(xdoc,xnode__header,c_xml_node__target,x)
  next

  dim xnode__data
  set xnode__data=xml_add_element(xdoc,xnode__install_files,c_xml_node__data)

  for each x in m_all_dirs__by_id.items
   call xnode__data.appendChild(x.create_xml_entry(xdoc))
  next

  for each x in m_all_files__by_id.items
   call xnode__data.appendChild(x.create_xml_entry(xdoc))
  next

  'formatting -------------------------------------------------------
  dim rdr
  set rdr=createobject("Msxml2.SAXXMLReader")

  dim wrt
  set wrt=createobject("Msxml2.MXXMLWriter")

  wrt.indent=true
  wrt.omitXMLDeclaration=true

  set rdr.contentHandler = wrt
  set rdr.dtdHandler = wrt
  set rdr.errorHandler = wrt

  rdr.parse xdoc

  'write to file ----------------------------------------------------
  dim stream
  set stream=createobject("ADODB.Stream")

  stream.Charset="utf-8"
  stream.Open

  call stream.WriteText("<?xml version=""1.0"" encoding=""utf-8""?>",1)
  call stream.WriteText(wrt.output,0)

  call stream.SaveToFile(file_path,2)
 end sub 'save_to_file

 '------------------------------------------------------------------------
 private sub add_include(include_name)
  if(m_includes.Exists(include_name))then
   call err.raise("Include ["&include_name&"] already exists")
  end if

  call m_includes.add(include_name,empty)
 end sub 'add_include

 '------------------------------------------------------------------------
 private sub add_target(target_name)
  if(m_targets.Exists(target_name))then
   call err.raise("Target ["&target_name&"] already exists")
  end if

  call m_targets.add(target_name,empty)
 end sub 'add_target

 '------------------------------------------------------------------------
 public function can_include(name)
  if(m_includes.Count=0)then
   can_include=true
   exit function
  end if

  dim x

  for each x in m_includes.keys
   if(ucase(name)=ucase(x))then
    can_include=true
    exit function
   end if
  next 'x

  can_include=false
 end function 'can_include

 '------------------------------------------------------------------------
 public sub reset_exists()
  dim x
  for each x in m_all_dirs__by_id.items
   call x.reset_exists()
  next

  for each x in m_all_files__by_id.items
   call x.reset_exists()
  next
 end sub 'reset_exists

 '------------------------------------------------------------------------
 public function reg_dir(parent_dir_id,name)
  if(IsNull(parent_dir_id))then
   call err.raise(-1,"t_install_files_database::reg_dir","Null parent_dir_id")
  end if

  if(not m_all_dirs__by_id.Exists(parent_dir_id))then
   call err.raise(-1,"t_install_files_database::reg_dir","Unknown parent_dir_id: "&parent_dir_id)
  end if

  dim dirs
  set dirs=m_all_dirs__by_id.Item(parent_dir_id).m_dirs

  if(dirs.exists(name))then
   set reg_dir=dirs.item(name)

   call reg_dir.set_exists()

   exit function
  end if

  dim dir
  set dir=new t_install_files_database__dir

  call dir.init(parent_dir_id,m_gen_dir_id.gen_id(),name)

  if(m_all_dirs__by_id.Exists(dir.id))then
   call err.raise(-1,,"Dublicate dir_id: "&dir.id)
  end if

  call dirs.add(name,dir)
  call m_all_dirs__by_id.add(dir.id,dir)

  set reg_dir=dir
 end function 'reg_dir

 '------------------------------------------------------------------------
 public function reg_file(parent_dir_id,name)
  if(IsNull(parent_dir_id))then
   call err.raise(-1,"t_install_files_database::reg_file","Null parent_dir_id")
  end if

  if(not m_all_dirs__by_id.Exists(parent_dir_id))then
   call err.raise(-1,"t_install_files_database::reg_file","Unknown parent_dir_id: "&parent_dir_id)
  end if

  dim files
  set files=m_all_dirs__by_id.Item(parent_dir_id).m_files

  if(files.exists(name))then
   set reg_file=files.item(name)

   call reg_file.set_exists()

   exit function
  end if

  dim file
  set file=new t_install_files_database__file

  call file.init(parent_dir_id,m_gen_file_id.gen_id(),name)

  if(m_all_files__by_id.Exists(file.id))then
   call err.raise("Dublicate file_id: "&file.id)
  end if

  call files.add(name,file)
  call m_all_files__by_id.add(file.id,file)

  set reg_file=file
 end function 'reg_file
end class 't_install_files_database

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
