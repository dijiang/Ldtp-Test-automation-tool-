#!/usr/bin/env python
# -*- coding:utf-8 -*-

# define some values

NORMAL=0
PTAB=1
LST=2
MNU_ITEM=3
TBL=4

special_dlg={   'btn文件':'*Word',
		'btn行和段落间距':'*Word',
		'btn背景颜色':'*Word',
		'btn更改大小写':'*Word'}

special_updlg={ 'btn编号开/关':'*Word',
		'btn项目符号开/关':'*Word',
		'btn下划线':'*Word',
		'btn粘贴':'*Word'}

mnu_item={	'mnu新建(N)':['mnu从模板新建(T)','mnu空白文档(D)'],
		'mnu最近的文档(U)':['mnu清除列表'],
		'mnu模板管理(T)':['mnu另存为模板(A)','mnu模板管理(B)']}

ls_type=['ptab','ptl','tbl','cbo','sbtn','dlg','rbtn','tree','ttbl','pane',
	'ico','frm','cal','pnl','lbl','mbr','mnu','lst','btn','tbtn','scbr'
	,'scpn','txt','auto','stat','tch','spr','flr','cnvs','splt','sldr',
	'html','pbar','tbar','ttip','chk','tblc','opane','popmnu','emb',
	'unk'	
	]

type_key={      'btn':[NORMAL,'btn'],
		'rbt':[NORMAL,'rbtn'],
		'tbt':[NORMAL,'tbtn'],
		'pta':[PTAB  ,'ptab'],
		'chk':[NORMAL,'chk'],
		'cbo':[NORMAL,'cbo'],
		'txt':[NORMAL,'txt'],
		'tbl':[TBL   ,'tbl']
		}	
