#!/usr/bin/env python
# -*- coding:utf-8 -*-
import ldtp
import val.define

class wrap:
    def __init__( self ) :
        self.appname = "*Word"
	self.special_dlg={'btn文件':'*Word','btn行和段落间距':'*Word','btn背景颜色':'*Word','btn更改大小写':'*Word'}
	self.special_updlg={'btn编号开/关':'*Word','btn项目符号开/关':'*Word','btn下划线':'*Word','btn粘贴':'*Word'}
        self.eventmap = { "btn" : [   'click( <parent>, <self> )', 
		                      'mouseleftclick( <parent>, <self> )', 
			              'stateenabled( <parent>, <self> )' ],
			  "tbtn": [   'mouseleftclick( <parent>, <self>)'],
			  "ptab": [   'selecttabindex( <parent>,<tab_list_name>,<self_index> )'],
			  "lst"	: [   'comboselectindex(<parent>,<component_name>,<self_index>)'],
			  "lst1": [   'mouseleftclick(<parent>,<self> )'],		
			  "rbtn": [   'click( <parent>, <self> )'],
			  "chk" : [   'click( <parent>, <self> )'],
			  "cbo" : [   'getallitem(<parent>, <self>)'],
			  "txt" : [   'settextvalue(<parent>,<self>,\'txt_write_by_yourself\')'],
			  "mnu" : [   'click( <parent>, <self> )'],
			  "mnui": [   'selectmenuitem(<parent>,<mnu;mnuitem>)'],
			  "tbl" : [   'selectrowindex(<parent>,<self>,<child_index>)',
				      'getrowcount(<parent>,<self>)',
				      'doubleclickrow(<parent>,<self>,\'column name\')']
		         }
	self.mnu_item={'mnu新建(N)':['mnu从模板新建(T)','mnu空白文档(D)'],
		       'mnu最近的文档(U)':['mnu清除列表'],
		       'mnu模板管理(T)':['mnu另存为模板(A)','mnu模板管理(B)']
			}
    
    def to_script( self, val ,flag ,key=None):
	print flag
	if(0==flag):
		val = val.replace( "<parent>", "\"%s\""%self.parent )
        	val = val.replace( "<self>", "\"%s\""%self.self )
	if(1==flag):  #ptab flag=1
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<tab_list_name>","\"ptl0\"")
		tmp=ldtp.getobjectproperty(self.parent,self.self,'child_index')
		val=val.replace("<self_index>","%d"%tmp)
	if(2==flag):  #lst flag=2
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<component_name>","\"%s\""%key)
		tmp=ldtp.getobjectproperty(self.parent,self.self,'child_index')
		val=val.replace("<self_index>","%d"%tmp)
	if(3==flag):  #mnu_item flag=3
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<mnu;mnuitem>","\"%s\""%key)
	if(4==flag):  #tbl flag=4
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<self>","\"%s\""%self.self)
		tmp=ldtp.getrowcount(self.parent,self.self)
		val=val.replace("<child_index>","0~%d"%(tmp-1))
        return val

    def get_type( self, name ):
        if 0== name.find("btn"):
		self.FLAG=0
		return "btn"
	if 0== name.find("rbtn"):
		self.FLAG=0
		return "rbtn"
	if 0== name.find("tbtn"):
		self.FLAF=0
		return "tbtn"
	if 0== name.find("ptab"):
		self.FLAG=1
		return "ptab"
	if 0== name.find("chk"):
		self.FLAG=0
		return "chk"
	if 0== name.find("cbo"):
		self.FLAG=0
		return "cbo"
	if 0== name.find("txt"):
		self.FLAG=0
		return "txt"
	if 0==name.find("tbl"):
		self.FLAG=4
		return "tbl"
	if 0== name.find("mnu"):
		component=[k for k,v in self.mnu_item.iteritems() if name in v]
		if component:
			self.FLAG=3    #mnu_item 子菜单
			return "mnui"
		else:
			self.FLAG=0    #单个mnu，和flag 0 处理一样
			return "mnu"
	if 0== name.find("lst"):
		component=[k for k,v in val.define.lstmap.iteritems() if name  in v]
		if component:
			component_name=component[0]
			print component_name
			if(component_name=="colour" or component_name=="notneed_parent"):  #lst color and no parents 特殊，和flag 0处理一样
				print 'match color or no parents'
				self.FLAG=0 
				return "lst1"
			else:
				self.FLAG=2
				return 'lst'
		else:
			print "kongde"
			self.FLAG=2
		return "lst"
        return None

    def event( self, name ): 
        self.self = name
        t = self.get_type( name )
        if None == t :
            return False

        self.ls = self.eventmap[ t ]
        return True

    def run( self ):
        ldtp.launchapp( "liteword" )
    
    def to_list(self,name):
	# 查找类似"btn文件"的特殊子窗口，需先关闭软件，不然找不到
	ldtp.launchapp('liteword')
	lst1=ldtp.getwindowlist()
	ldtp.mouseleftclick(self.special_dlg.get(name),name)
	lst2=ldtp.getwindowlist()
	ret=list(set(lst1)^set(lst2))
	print 'hehe :',ret
	if ret:
		for i in ret:
			if i.startswith("dlg"):
				return i
    def to_uplist(self,name):
	# 查找类似"btn文件"的特殊子窗口，需先关闭软件，不然找不到
	ldtp.launchapp('liteword')
	lst1=ldtp.getwindowlist()
	ldtp.click(self.special_updlg.get(name),name)
	ldtp.generatekeyevent('<up>')
	lst2=ldtp.getwindowlist()
	ret=list(set(lst1)^set(lst2))
	print 'hehe :',ret
	if ret:
		for i in ret:
			if i.startswith("dlg"):
				return i

         

    def list( self, name ):
	#先判断是不是特殊的，如果是就修改成真正要查找的名字 
	self.parent = None
	if(name in self.special_dlg.keys()):
		name=self.to_list(name)
	if(name in self.special_updlg.keys()):
	    	name=self.to_uplist(name)
	print name
        try :
  	    if 0 == len( name ):
                self.ls = ldtp.getwindowlist()
            elif 1 == ldtp.guiexist( name ):
                self.ls = ldtp.getobjectlist( name )
            else:
                self.msg = "没有找到窗口"
                return False
        except :
            self.msg = "没有找到窗口"
            return False

        v = len( self.ls )
        if 0 == v:
            self.msg = "没有找到窗口"
            return False
     
        self.msg = "找到" + "%i"%v + "个窗口"
        self.parent = name
        return True
	
