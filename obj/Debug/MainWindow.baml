   M S B A M L   `   `   `  ����M  IEstadoResultadoRCL, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:EstadoResultadoRCLEstadoResultadoRCL  5   	   #    EstadoResultadoRCL.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   2local clr-namespace:EstadoResultadoRCL     5   	     � Title$'  !Procesamiento de Estado Resultado��5   	   $��540.594��5	   	   $��663.97��5
   	   .��5      � $��0,0,20,7q�6   .+�5   
   o� $	��337��5      $��0��5      $��0��5      $
��Left=�5      $��Stretch=�5      $��-4,-2,0,462.333333333333q�5      .@�5      m� -   5         m�Name$ mnData��5      $
��Data��5       J� RenderTransformOrigin$ 0.5,0.5&�5      .@�5      m� -   5        $ mnItem��5      $��Items��5      $
��Left=�5      5      m� 5      -   5        $ mnLineas��5      $!��Líneas de Estado Resultado��5      $
��Left=�5      m� 5       -   5        $ mnAreas��5!      $��Areas��5"      $
��Left=�5#      5$      5%      m� 5&      -   5        $ mnConfig��5'      $��Configuración��5(      $ 0.5,0.5&�5*      m� 5+      -   5        $ mnCtaAnalisis��5,      $��Cuenta análisis��5-      $ 0.5,0.5&�5/      m� 50      -   5        $ mnCtasCltes��51      $��Cuentas clientes��52      $ 0.5,0.5&�54      55      �� 56   
   -   5         ��Name$ PathIn��57      $
��Left=�58      $��28��59      $��
120,48,0,0q�5:       �� TextWrapping$
 Wrap=�5;      $	��Top=�5<      $	��490��5=      {�.  5>      �� 5?   
   -	   5         ��Name$ 	PathInBtn��5@      $��Ruta Entrada��5A      $
��Left=�5C      $��28��5D      $��	10,48,0,0q�5E      $	��Top=�5F      $��99��5G      �� 5H   
   -
   5        $ PathOut��5I      $
��Left=�5J      $��28��5K      $��
120,85,0,0q�5L      $
 Wrap=�5M      $	��Top=�5N      $	��490��5O      {�.  5P      �� 5Q   
   -   5        $ 
PathOutBtn��5R      $��Ruta Salida��5S      $
��Left=�5U      $��28��5V      $��	10,85,0,0q�5W      $	��Top=�5X      $��99��5Y      �� 5Z   
   -   5         ��Name$ ListInputFiles��5[      $
��Left=�5\      $��10,194,0,10q�5]      $	��443��5^       �� SelectionMode$ Multiple=�5_      �� 5`   
   -   5        $ ProcBtn��5b      $��Procesar��5a      $
��Left=�5d      $��32��5e      $��
10,158,0,0q�5f      $	��Top=�5g      $	��443��5h      �� 5i   
   -   5        $ FileOut��5j      $
��Left=�5k      $��28��5l      $��120,121,0,0q�5m      $
 Wrap=�5n      $	��Top=�5o      $	��490��5p      {�.  5q      �� 5r   
   $��Archivo de Salida��5s      $
��Left=�5t      $��28��5u      $��
10,121,0,0q�5v      $	��Top=�5w      $	��105��5x      �� 5y   
   -   5         ��Name$ chkApplyRate��5z      $��Aplicar tipo de cambio��5|      $
��Left=�5}      $	��Top=�5~      $��486,174,0,0q�5�      �� 5�   
   -   5        $ txtExchgRate��5�      $��Right=�5�      $��23��5�      $��
0,194,10,0q�5�      $
 Wrap=�5�      $	��Top=�5�      $	��139��5�      {�.  5�      5�      5�      