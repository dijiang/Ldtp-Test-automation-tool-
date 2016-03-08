#!/usr/bin/env python
# -*- coding:utf-8 -*-
import ldtp

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

	self.lstmap={"cbo字体名称":['lstBitstreamCharter','lstCenturySchoolbookL','lstCourier10Pitch','lstDejaVuSans','lstDejaVuSansCondensed','lstDejaVuSansLight','lstDejaVuSansMono','lstDejaVuSerif','lstDejaVuSerifCondensed','lstDingbats','lstDroidArabicNaskh','lstDroidSans','lstDroidSansArmenian','lstDroidSansEthiopic','lstDroidSansFallback','lstDroidSansGeorgian','lstDroidSansHebrew','lstDroidSansJapanese','lstDroidSansMono','lstDroidSansThai','lstDroidSerif','lstFreeMono','lstFreeSans','lstFreeSerif','lstLiberationMono','lstLiberationSans','lstLiberationSansNarrow','lstLiberationSerif','lstMTExtra','lstNimbusMonoL','lstNimbusRomanNo9L','lstNimbusSansL','lstOpenSymbol','lstStandardSymbolsL','lstSymbol','lstURWBookmanL','lstURWChanceryL','lstURWGothicL','lstURWPalladioL','lstWebdings','lstWingdings','lstWingdings2','lstWingdings3','lst文泉驿微米黑','lst文泉驿等宽微米黑'],

		     "cbo字体大小":['lst八号','lst七号','lst小六','lst六号','lst小五','lst五号','lst小四','lst四号','lst小三','lst三号','lst小二','lst二号','lst小一','lst一号','lst小初','lst初号','lst6','lst7','lst8','lst9', 'lst10','lst105','lst11','lst12','lst13','lst14','lst15','lst16','lst18','lst20','lst22','lst24','lst26','lst28','lst32','lst36','lst40','lst44','lst48','lst54','lst60','lst66','lst72','lst80','lst88','lst96'],
		
		     "cbo行距(L)":['lst单倍行距','lst15倍行距','lst双倍行距','lst按比例','lst至少','lst行间距离','lst固定'],			

		     "notneed_parent":['lst点线(粗)','lst点线','lst粗线','lst双线','lst单线','lst两点一划','lst波浪线','lst长虚线','lst(无)','lst(无)1','lst虚线','lst(无)','lst点划线','lstNumbering','lst编号对齐方式：左对齐6','lst编号对齐方式：左对齐7','lst编号对齐方式：左对齐4','lst编号对齐方式：左对齐5','lst编号对齐方式：左对齐2','lst编号对齐方式：左对齐3','lst无','lst编号对齐方式：左对齐1','lst编号对齐方式：左对齐','lst核对符号项目符号','lst实心大圆形项目符号','lstBullet','lst绿色圆形','lst实心菱形项目符号','lst右指箭头项目符号','lst无','lst实心大正方形项目符号','lst右指箭头项目符号1','lst红色方块','lst行距15倍','lst行距(L)','lst行距双倍','lst行距115倍','lstLineSpacing','lst行距单倍','lst行距单倍1'],

	   	     "colour":['lst自动','lst黑色','lst白色','lst灰色1','lst灰色2','lst灰色3','lst灰色4','lst灰色5','lst灰色6','lst灰色7','lst灰色8','lst灰色9','lst灰色10','lst黄色','lst橙色','lst红色','lst粉红','lst紫红色','lst紫色','lst蓝色','lst天蓝','lst青色','lst蓝绿色','lst绿色','lst黄绿','lst黄色1','lst橙色1','lst红色1','lst粉红1','lst紫红色1','lst紫色1','lst蓝色1','lst天蓝1','lst青色1','lst蓝绿色1','lst绿色1','lst黄绿1','lst黄色2','lst橙色2','lst红色2','lst粉红2','lst紫红色2','lst紫色2','lst蓝色2','lst天蓝2','lst青色2','lst蓝绿色2','lst绿色2','lst黄绿2','lst黄色3','lst橙色3','lst红色3','lst粉红3','lst紫红色3','lst紫色3','lst蓝色3','lst天蓝3','lst青色3','lst蓝绿色3','lst绿色3','lst黄绿3','lst黄色4','lst橙色4','lst红色4','lst粉红4','lst紫红色4','lst紫色4','lst蓝色4','lst天蓝4','lst青色4','lst蓝绿色4','lst绿色4','lst黄绿4','lst黄色5','lst橙色5','lst红色5','lst粉红5','lst紫红色5','lst紫色5','lst蓝色5','lst天蓝5','lst青色5','lst蓝绿色5','lst绿色5','lst黄绿5','lst黄色6','lst橙色6','lst红色6','lst粉红6','lst紫红色6','lst紫色6','lst蓝色6','lst天蓝6','lst青色6','lst蓝绿色6','lst绿色6','lst黄绿6','lst黄色7','lst橙色7','lst红色7','lst粉红7','lst紫红色7','lst紫色7','lst蓝色7','lst天蓝7','lst青色7','lst蓝绿色7','lst绿色7','lst黄绿7','lst黄色8','lst橙色8','lst红色8','lst粉红8','lst紫红色8','lst紫色8','lst蓝色8','lst天蓝8','lst青色8','lst蓝绿色8','lst绿色8','lst黄绿8','lst黄色9','lst橙色9','lst红色9','lst粉红9','lst紫红色9','lst紫色9','lst蓝色9','lst天蓝9','lst青色9','lst蓝绿色9','lst绿色9','lst黄绿9','lst黄色10','lst橙色10','lst红色10','lst粉红10','lst紫红色10','lst紫色10','lst蓝色10','lst天蓝10','lst青色10','lst蓝绿色10','lst绿色10','lst黄绿10','lst蓝灰色','lst经典蓝','lst深红色','lst浅黄色','lst淡绿色','lst深紫色','lst橙红色','lst湖蓝色','lst图表1','lst图表2','lst图表3','lst图表4','lst图表5','lst图表6','lst图表7','lst图表8','lst图表9','lst图表10','lst图表11','lst图表12','lst探戈天蓝1','lst探戈天蓝2']		
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
		component=[k for k,v in self.lstmap.iteritems() if name  in v]
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
	
