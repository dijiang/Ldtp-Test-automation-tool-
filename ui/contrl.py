# -*- coding: utf-8 -*-
#!/usr/bin/env python
import gtk

class tkView():
    def __init__( self ) :
	self.win = gtk.TextView()
	self.win.show()
        self.buf = self.win.get_buffer()


    def insert( self, val ):
        self.buf.insert_at_cursor( val + "\n" )

class tkList():
    def __init__( self, name, clickfun = None ) :   #name=控件消息/窗口名称
        self.clickfun = clickfun

        self.listchild = gtk.ListStore( str )
        self.listchild.set_default_sort_func( None )
        self.tree = gtk.TreeView( self.listchild )
        col = gtk.TreeViewColumn( name )
        self.tree.append_column( col )
        cell = gtk.CellRendererText()
        col.pack_start( cell, True )
        col.add_attribute( cell, 'text', 0 )
        self.tree.show()

        self.win = gtk.ScrolledWindow()
        self.win.set_border_width( 10 )
        self.win.set_size_request( 300, 300 )
        self.win.set_policy( gtk.POLICY_AUTOMATIC, gtk.POLICY_ALWAYS )
        self.win.add_with_viewport( self.tree )
        self.win.show()

    def setDoubleClick( self, fun ):
        self.tree.connect( "row-activated", fun, None )

    def clear( self ):
        self.listchild.clear()
	
	#控制窗口名称下面打印的内容
    def append( self, ls ):
        mod = self.tree.get_model()
        self.tree.set_model( None )
	ls=sorted(ls)
        for s in ls :
		#过滤
		#if s.startswith('flr') or s.startswith('frm') or s.startswith('pnl') or s.startswith('tbar') or s.startswith('ukn') or s.startswith('lbl'):
		#	pass
		#else:
		self.listchild.append( [ s ] )

        self.tree.set_model( mod )

    def cur_val( self, path ):
        mod = self.tree.get_model()
        it = mod.get_iter( path )
        val = mod.get_value( it, 0 )
        return val

