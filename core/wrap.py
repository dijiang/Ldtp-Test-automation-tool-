#!/usr/bin/env python
# -*- coding:utf-8 -*-
import ldtp
import define
import os
import re

FILE_PATH=os.path.expanduser('~')+'/typedef.txt'
INIT_PATH='./.init.txt'
class wrap:
    def __init__( self ) :
        self.appname = "*Word"
	if not(os.path.exists(INIT_PATH)):
		with open(INIT_PATH,'w'):
			pass
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
    
    def to_script( self, val ,flag ,key=None):
	print flag
	if(define.NORMAL==flag):
		val = val.replace( "<parent>", "\"%s\""%self.parent )
        	val = val.replace( "<self>", "\"%s\""%self.self )
	if(define.PTAB==flag):  #ptab flag=1
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<tab_list_name>","\"ptl0\"")
		tmp=ldtp.getobjectproperty(self.parent,self.self,'child_index')
		val=val.replace("<self_index>","%d"%tmp)
	if(define.LST==flag):  #lst flag=2
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<component_name>","\"%s\""%key)
		tmp=ldtp.getobjectproperty(self.parent,self.self,'child_index')
		val=val.replace("<self_index>","%d"%tmp)
	if(define.MNU_ITEM==flag):  #mnu_item flag=3
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<mnu;mnuitem>","\"%s\""%key)
	if(define.TBL==flag):  #tbl flag=4
		val=val.replace("<parent>","\"%s\""%self.parent)
		val=val.replace("<self>","\"%s\""%self.self)
		tmp=ldtp.getrowcount(self.parent,self.self)
		val=val.replace("<child_index>","0~%d"%(tmp-1))
        return val

    def get_type( self, name ):	
	key=name[:3]
	if(key in define.type_key.iterkeys()):
		self.FLAG=define.type_key[key][0]
		return define.type_key[key][1]
	if (name.startswith("mnu")):
		component=[k for k,v in define.mnu_item.iteritems() if name in v]
		if component:
			self.FLAG=3    #mnu_item 子菜单
			return "mnui"
		else:
			self.FLAG=0    #单个mnu，和flag 0 处理一样
			return "mnu"
	if (name.startswith("lst")):
		component=[k for k,v in self.lst_map.iteritems() if name  in v]
		if component:
			component_name=component[0]  #lst 属于special cbo,flag=2 处理
			print component_name
			self.FLAG=2
			return "lst"
		else:				     #其他lst,flag=o 默认处理
			self.FLAG=0
			return "lst1"
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

	#处理别名的函数:
    def typedef(self,ls):
	ls=sorted(ls)
	with open(INIT_PATH,'r+') as init:
		self.total_file=[i.strip('\n') for i in init.readlines()]
	n=0
	with open(FILE_PATH,'a+') as f:
		for i in ls:
			if i not in self.total_file:
				for name in define.ls_type:
					if(i.startswith(name)):
						n=n+1
						a=i[len(name):]
						f.write("%s:%s\n"%(i,a))
						with open(INIT_PATH,'a+') as end:
							end.write('%s\n'%i)
		if(n>0):
			f.write("**************************%d*****************************\n"%n)
    
	#处理lst分类的函数:
    def cope_lst(self,name):
	lst_map={}
	cbo=[i for i in self.ls if i.startswith('cbo')]
	for i in cbo:
		lst_tmp=[]
		for tmp in ldtp.getallitem(name,i):
			count=0
			newstr='lst'+''.join(re.split('\s+|\.','%s'%tmp))
			#判断不同的cbo中是否有同名的lst,如果存在按照ldtp规则，在后面依次加入1,2...
			for n in lst_map.values():
				for item in n:
					if (newstr==item or newstr==item[:len(n)-1]):
						count+=1
			if(count>0):
				lst_tmp.append(newstr+'%d'%count)
			else:
				lst_tmp.append(newstr)
		lst_map[i]=lst_tmp		
	#print "lst_map:",lst_map
	return lst_map

    def to_list(self,name):
	# 查找类似"btn文件"的特殊子窗口，需先关闭软件，不然找不到
	ldtp.launchapp('liteword')
	lst1=ldtp.getwindowlist()
	ldtp.mouseleftclick(define.special_dlg.get(name),name)
	lst2=ldtp.getwindowlist()
	ret=list(set(lst1)^set(lst2))
	print 'hehe :',ret
	if ret:
		for i in ret:
			if i.startswith("dlg"):
				return i
    def to_uplist(self,name):
	# 查找类似"btn下划线"的特殊子窗口，需先关闭软件，不然找不到
	ldtp.launchapp('liteword')
	lst1=ldtp.getwindowlist()
	ldtp.click(define.special_updlg.get(name),name)
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
	self.lst_map={}
	if(name in define.special_dlg.keys()):
		name=self.to_list(name)
	if(name in define.special_updlg.keys()):
	    	name=self.to_uplist(name)
	print name
        try :
  	    if 0 == len( name ):
                self.ls = ldtp.getwindowlist()
            elif 1 == ldtp.guiexist( name ):
                self.ls = ldtp.getobjectlist( name )
		self.lst_map=self.cope_lst(name)
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
	
