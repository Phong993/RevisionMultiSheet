   M S B A M L   `   `   `  ���� N  JMultiSelectComboBox, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullV>clr-namespace:MultiSelectComboBox;assembly=MultiSelectComboBoxMultiSelectComboBox  P LRevisionMultiSheet, Version=2017.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:RevisionMultiSheetRevisionMultiSheet #   RevisionMultiSheet.FormMaster   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     Rcontrol>clr-namespace:MultiSelectComboBox;assembly=MultiSelectComboBox     2local clr-namespace:RevisionMultiSheet     -   $	��724��$	��900��$
��Left=�$
��Left=�$	��Top=�$	��Top=�.��� -     �Name$
  grid��$��0,0,0,0q�$
��Left=�$	��890��$	��724��$	��Top=� � ColumnDefinitions �� $	��44* ��� $	��45* � � RowDefinitions �� $
��Auto ��� $
��Auto �$	��313���� $
��Auto �$��32���� $
��Auto �$	��353���� $
��Auto �.+��� $��Revision List��$��Verdana)�$
��Left=�$��Bottom=�$��10,0,0,0q�$��0��$��0��$��25��$��85���� $3��-Note: Right Click on list Revision to Delete.��$��Verdana)�$
��Left=�$��Bottom=�$��	150,0,0,0q�$��0��$��0��$��25��$	��290��� -    �Name$ progressBar��$	��400��$��10��$
��Left=�$	��Top=�$��	10,30,0,0q�$��1��$��1���� -    ��Name$ 	resultBox��$
��Left=�$	��Top=�$��	10,50,0,0q�$��1��$��1��$	��400��$	��250��&   System.Windows.Controls.DataGrid  -     Name$ gridRevision��   SelectionMode$ Single=�$	��400��$	��300��$��10,0,0,0q�$
��Left=�$	��Top=�$��1��   
IsReadOnly .  � ContextMenu �� .@�m� -   $��Delete This Revision��m� -   $��Delete All Revision���� $
��Left=�$	��Top=�$��9,8,0,0q�$��2��$	��400��$��0�����	#FF959494�� $
��Left=�$	��Top=�$��9,8,0,0q�$��2��$	��400��$��1�����	#FF959494�� $
��Left=�$	��Top=�$��347,218,0,0q�$��1��$	��206�����	#FF959494	 J� RenderTransformOrigin$	 0.5,0.5&�$��2��z�V� .~��� �� �� 
 �� Angle$
 90.424Q�T� �� $��Already Sheets��$��Verdana)�$
��Left=�$��Bottom=�$��10,0,0,0q�$��2��$��0��$��25��$	��101��  -   $ alreadySheet��$	��400��$	��300��$��	10,30,0,0q�$
��Left=�$	��Top=�$��3�� .    Columns  �� .@�m� -	    m�Name$! alreadySheet_Add_RightClick��$��Add All Sheet Selected��m� -
   $ alreadySheet_RightClick��$��Remove All Revisions��m� -   $. (alreadySheet_RemoveCurrentRev_RightClick��$��Remove Current Revision��m� -   $��Go to View...���� $��Selected Sheets��$��Verdana)�$
��Left=�$��Bottom=�$��10,0,0,0q�$��2��$��1��$��25��$	��101��  -   $ selectedSheet��$	��400��$	��300��$��	10,30,0,0q�$
��Left=�$	��Top=�$��3��$��1�� . �� -    ��Name$ 	btnAccept��$��0q�$��60��$��20��$��Right=�$	��Top=�$��	0,0,200,0q�$��1��$��1��.����  �� Orientation$ 
Horizontal=�.+��� $��16��$��15��$
��Left=�$��0,1,5,0q�$��Center=�$"��image/button_start_16x18.png���� $��0q�.��OK�� -   $ 	btnCancel��$��0q�$��70��$��20��$��Right=�$	��Top=�$��	0,0,120,0q�$��1��$��1��.���� $ 
Horizontal=�.+��� $��16��$��15��$
��Left=�$��0,1,5,0q�$��Center=�$��image/cancel_16x16.png���� $��0q�.��Cancel�� -   $ 	btnReload��$��0q�$��70��$��20��$��Right=�$	��Top=�$��0,0,40,0q�$��1��$��1��.���� $ 
Horizontal=�.+��� $��16��$��15��$
��Left=�$��0,1,5,0q�$��Center=�$��image/reload_16x16.png���� $��0q�.��Reload�� $��Find Sheet:��$��Verdana)�$
��Left=�$	��Top=�$��10,0,0,0q�$��3��$��0��$��25��$��80���� -    ��Name$ 	findSheet��$	��150��$��20��$
��Left=�$	��Top=�$��90,2,0,0q�$��3��$��0��$��Find by Sheet Number���� -   $ btnAdd��$��Add >>��$
��Left=�$��350,0,0,328q�$��3��$��0��$��60��$��25���� -   $ 	btnRemove��$��Remove��$
��Left=�$	��Top=�$��10,0,0,0q�$��3��$��1��$��25��$��60���� $
��Left=�$��Bottom=�$��9,8,0,17q�$��3��$	��840��$��0��$��2�����	#FF959494�� $S��MHoa Binh Coporation - BIM Department - Revit Addins Team - version 2.0.0.2017��$��Verdana)�$��Center=�$��Bottom=�$��214,20,154,-7q�$��3��$��0��$��2��$��25��$	��522��$��2���� -    ��Name$ iconBIM��$��1��$��1��$��50��$��50��$��Right=�$��Bottom=�$��0,0,40,0q�-   'MultiSelectComboBox.MultiSelectComboBox  -     Name$ listCheckbox��$��70��$��25��$
��Left=�$��250,0,0,328q�$��3��$��0��