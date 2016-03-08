#!/usr/bin/env python
# -*- coding: utf-8 -*-
import pygtk
pygtk.require( '2.0' )
import gtk
from ui.contrl import *
import core.wrap
import define

class Application:
    def __init__( self ) :
        self.core = core.wrap.wrap()

        window = gtk.Window( gtk.WINDOW_TOPLEVEL )
        window.set_border_width( 10 )
        window.set_size_request( 1100, 600 )
        window.set_title( "LiteWord自动化测试工具" )
        window.connect( "delete_event", lambda w, e : gtk.main_quit() )
	self.old=[]
	#column
        hbox = gtk.HBox( False, 6 )
			 #   gtk.HBox(homogeneous=False, spacing=0)
			 #		homogeneous :
			 #			If True all children are given equal space allocations.
			 #		spacing :
			 #			The additional horizontal space between children in pixels
        window.add( hbox )
        hbox.show()
	
        hbox.pack_start( self.ui_window(), False, False )
        hbox.pack_start( self.ui_event(), False, False )
        hbox.pack_start( self.ui_view(), True, True )

        window.show()
	#最右边一栏
    def ui_view( self ):
        vbox = gtk.VBox( False, 3 )
        vbox.show()

        self.view = tkView()
		
        vbox.pack_start( self.view.win, True, True )

        return vbox
	
	#中间一栏
    def ui_event( self ):
        vbox = gtk.VBox( False, 3 )
        vbox.show()

	# row
        self.event_msg = gtk.Label( "控件ID: " )
        self.event_msg.show()
        vbox.pack_start( self.event_msg, False, False )

	# row
        self.event = tkList( "控件消息" )
        self.event.setDoubleClick( self.do_event )
        vbox.pack_start( self.event.win, True, True )

        return vbox

	#最左边一栏
    def ui_window( self ):
        vbox = gtk.VBox( False, 3 )
        vbox.show()

	# row
        w = gtk.Button( "运行liteword" )
        w.connect( "clicked", lambda w, e : self.core.run(), None )
        w.show()
        vbox.pack_start( w, False, False )

	# row
        hbox = gtk.HBox( False, 3 )

        w = gtk.Label( "窗口" )
        w.show()
        hbox.pack_start( w, False, False )

        self.winName = gtk.Entry()
        self.winName.show()
        self.winName.set_text( self.core.appname )
        hbox.pack_start( self.winName, False, False )

        w = gtk.Button( "查找子窗口" )
        w.connect( "clicked", self.fun_search,None)
        w.show()
        hbox.pack_start( w, False, False )
	
	w = gtk.Button( "返回" )
	w.connect( "clicked", self.fun_research, None )
	w.show()
	hbox.pack_start( w, False, False )

        hbox.show()
        vbox.pack_start( hbox, False, False )

	# row
        self.message = gtk.Label( "" )
        self.message.show()
        vbox.pack_start( self.message, False, False )

	# row
        self.tree = tkList( "窗口名称" )
        self.tree.setDoubleClick( self.do_control )
        vbox.pack_start( self.tree.win, True, True )

        return vbox
	#控件消息双击执行的函数
    def do_event( self, widget, path, col, data = None ):
        val = self.event.cur_val( path )
	#print flag
	text=self.winName.get_text()

	if(2==self.core.FLAG):   #处理flag=2（lst）的特殊情况
		component=[k for k,v in define.lstmap.iteritems() if text in v]
		if component:
			component_name=component[0]
		else:
			component_name=None
			#return
		val = self.core.to_script( val,self.core.FLAG,component_name )
		self.view.insert( val )
	elif(3==self.core.FLAG): #处理flag=3(mnu_item)的特殊情况
		component=[k for k,v in self.core.mnu_item.iteritems() if text in v]
		if component:
			component_name=component[0]
		else:
			component_name=None
			return
		key=component_name+';'+text[3:]
		print 'key: '+key
		val =self.core.to_script(val,self.core.FLAG,key)
		self.view.insert(val)
	else:			#处理其他通用情况
		val = self.core.to_script( val, self.core.FLAG)
		self.view.insert( val )
		

	#窗口名称下双击执行的函数
    def do_control( self, widget, path, col, data = None ):
        val = self.tree.cur_val( path )
        self.winName.set_text( val )                  #更改窗口旁边的内容

        self.event_msg.set_text( "控件ID: %s"%val )   #更改控件ID的标题
        self.event.clear()
        if self.core.event( val ):
            self.event.append( self.core.ls )      #中间一栏显示的控件消息

	#查找子窗口按钮执行函数：
    def fun_search( self, widget, data  ):      
        self.tree.clear()

        name = self.winName.get_text()
	if self.old:
		if(name!=self.old[-1]):
	    		self.old.append(name)
	else:
		self.old.append(name)
        if self.core.list( name ):
            self.tree.append( self.core.ls )        #窗口名称下面显示找到的东西
	self.message.set_text( self.core.msg )  #按钮下面显示找到270个窗口
	#返回按钮执行函数：
    def fun_research( self, widget, data = None ):
	self.tree.clear()
	print 'hiahia',self.old
	if self.old:
		self.old.pop()
		if self.old:
			if self.core.list(self.old[-1]):
				self.tree.append( self.core.ls )        #窗口名称下面显示找到的东西
			self.message.set_text( self.core.msg )  #按钮下面显示找到270个窗口
			self.winName.set_text( self.old[-1] )   #更改窗口旁边的内容	
		else:
			self.old.append(self.core.appname)
			self.message.set_text("")

    def main( self ):
        gtk.main()
	
    def mytype(self,name):
	if(0==name.find('ptab')):
		return 1
	if(0==name.find('lst')):
		return 2
	return 0

if __name__ == "__main__" :
    app = Application()
    app.main()
