   M S B A M L   `   `   `  ����   GridWpf    GridWpf.UserControl1   \ XPresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    8x,http://schemas.microsoft.com/winfx/2006/xaml    Hmc;http://schemas.openxmlformats.org/markup-compatibility/2006    >d2http://schemas.microsoft.com/expression/blend/2008      � 	Resources     HeaderBackground �� EndPoint �� 
StartPoint � Color � Offset  MyBeckgrHeader2��%L  &	        &	 �     �� $ 0.5,1&�$ 0.5,0&�.��� $ 	#FF4FC8D4��$ 0.547Q�� $ 	#FFE1FFFF��$ 1Q�� $ 	#FFE2FFFF��$ 0Q�� $ 	#774FC8D4��$ 0.797Q�� $ 	#774FC8D4��$ 0.291Q��� $ 0.5,1&�$ 0.5,0&�.��� $ 	#FF00E8E8��$ 0Q�� $ 	#FF2FB7E3��$ 0.523Q�� $ 	#FF00E8E8��$ 1Q�.��� $	��308��$
��1344�� � ColumnDefinitions �� $	��59* ��� $	��56* ��� $
��208* ��� $
��145* ��� $	��92* ��� $	��80* ��� $	��86* ��� $	��71* ��� $
��111* ��� $	��67* ��� $	��78* ��� $
��101* ��� $	��79* ��� $	��85* ��� $	��83* � � RowDefinitions �� $	��53* ��� $	��32* ��� $
��113* ��� $
��161* �.+��� -   $��border2��
�������$��1q�$��2��$��3������  �� Source# [  ._��� -   $��
textBlock1��$��38��$6��0Наименование  мероприятия��$��WrapWithOverflow=�$	��147��$��Stretch=�$��Center=� �� TextAlignment$ Center=��� -   $��border3��
�������$��1q�$��1��$��3������ # [  ._��� -   $��
textBlock3��$��93��$6��0Номер меро-приятия в плане��$��WrapWithOverflow=�$ Center=�$��Center=�$��Center=��� -   $��border4��
�������$��1q�$��3������ # [  ._��� -   $��
textBlock2��$	��138��$V��PКод  основных направлений энергосбережения��$ Center=�$��Center=�$��Center=�$��49��$
��Wrap=��� -   $��border5��
�������$��1q�$��3��$��3������ # [  ._��� -   $��
textBlock4��$��42��$!��Дата внедрения��$ Center=�$��WrapWithOverflow=�$��Center=��� -	   $��border6��$��1q�$��4��$��2��$��2������ # [  ._��� -
   $��
textBlock5��$��25��$#��Объем внедрения��$��Center=�$��WrapWithOverflow=�$ Center=��� -   $��border7��$��1q�$��4��$��2������ # [  ._��� -   $��
textBlock6��$��37��$'��!Единица измерения��$��Center=�$ Center=�$��WrapWithOverflow=��� -   $��border8��$��1q�$��6��$��3������ # [  ._��� -   $��
textBlock8��$��53��$4��.Экономия ТЭР, т. усл. топл.��$��Center=�$��WrapWithOverflow=�$ Center=��� -   $��border1��
�������$��1q�$��8��$��7��$��1������ # [  ._��� -   $��textBlock11��$��23��$=��7из них за счет средств бюджета��$��Center=�$��Center=�$��WrapWithOverflow=�$ Center=��� -   $��border10��
�������$��1q�$��7��$��8��$��1,0,0,0q����� # [  ._��� -   $��
textBlock9��$��20��$����Затраты на внедрение мероприятия  и источники финансирования, млн. руб.��$��Center=�$��WrapWithOverflow=�$ Center=��� -   $��border11��
�������$��1q�$��7��$��1��$��2������ # [  ._��� -   $��textBlock10��$��23��$��
Всего��$��Center=�$ Center=�$��Center=�$��WrapWithOverflow=��� -   $��border12��
�������$��1q�$��8��$��2������ # [  ._��� -   $��textBlock12��$��93��$r��lресп. бюджета на финансирование программ энергосбережения��$��Center=�$ Center=�$��Stretch=�$
��Wrap=�$��84���� -   $��border13��
�������$��1q�$��9��$��2������ # [  ._��� -   $��textBlock13��$��57��$/��)инвестиционных фондов��$��Center=�$
��Wrap=�$ Center=�$��Center=��� -   $��border14��
�������$��1q�$��10��$��2������ # [  ._��� -   $��textBlock14��$��59��$5��/республиканского бюджета��$��Center=�$ Center=�$
��Wrap=��� -   $��border15��
�������$��1q�$��11��$��2������ # [  ._��� -   $��textBlock15��$��34��$%��местного бюджета��$��Center=�$
��Wrap=�$ Center=��� -   $��border16��
�������$��1q�$��12��$��2������ # [  ._��� -   $��textBlock16��$��25��$��организации��$��Center=��� -   $��border17��
�������$��1q�$��13��$��2������ # [  ._��� -    $��textBlock17��$��23��$��кредита��$��Center=�$ Center=�$
��Wrap=��� -!   $��border18��
�������$��1q�$��14��$��2������ # [  ._��� -"   $��textBlock18��$��48��$'��!Других источников��$��Center=�$
��Wrap=�$ Center=�$��53���� -#   $��gridSplitter1��$��5��$��2��$	��113��$	��Top=�$��0,0,1,0q����� # [  �� -$   $��gridSplitter2��$��5��$��1��$��0,30,1,0q�$��2������ # [  �� -%   $��border9��$��1q�$��5��$��2��$��0,0,1,0q����� # [  ._��� -&   $��
textBlock7��$��Количество��$��46��$
��Wrap=�$ Center=�	 �� IsHyphenationEnabled	 .  $��36��