   M S B A M L   `   `   `  ����   ExportDB    ExportDB.MainWindow   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml       � Title$  ExportDB��$	��400��$	��600��.���  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+��� $��0��$
��Left=�$��Center=�$��	10,10,0,0q��� $�DB导出路径: �� �� -    ��Name$ DBExportPath���� -   $��0��$��选择DB导出路径��$��Right=�$��Center=�$��	0,10,20,0q�$	��135���� -   $��1��$��添加Excel文件��$
��Left=�$��Center=�$��	10,10,0,0q��� -    ��Name$ FileList��$��2��$��Stretch=�$��Stretch=�$��10,10,10,10q�����  �� 
TargetType" ��.���  �� Property � �� �� Value$ Stretch=����� .R��  � ColumnDefinitions �� $
��Auto ��� $��* ��� $
��Auto �.+���  �� Orientation$ 
Horizontal=�$��0��$��0,5,0,5q�.+��� $��16��$��16��$��0,0,10,0q����� *Icon+�� $��Center=����� *DisplayName+�� -   $��2��$��	10,0,10,0q����� 	 �� 
TargetType"	 ��.R�� $��20��$��20��
���  ��.+��� $��X��$��Center=�$��Center=�
��������� -   $��3��$��开始导出��$��Stretch=�$��Center=�$��10,20,10,10q�
��� � �$��40��