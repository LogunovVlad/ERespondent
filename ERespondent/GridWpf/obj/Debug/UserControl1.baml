   M S B A M L   `   `   `  ����  GridWpf    GridWpf.UserControl1   5      \ XPresentationFramework, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      8x,http://schemas.microsoft.com/winfx/2006/xaml    5      Hmc;http://schemas.openxmlformats.org/markup-compatibility/2006    5      >d2http://schemas.microsoft.com/expression/blend/2008    5        � 	Resources  5         HeaderBackground �� EndPoint �� 
StartPoint � Color � Offset  MyBeckgrHeader2��%  &	        &	 8    �� 5   
   $ 0.5,1&�67   $ 0.5,0&�6H   .��5	      � $ 	#FF4FC8D4��6   $ 0.547Q�6-   � 5
      $ 	#FFE1FFFF��6   $ 1Q�6-   � 5      $ 	#FFE2FFFF��6   $ 0Q�6-   � 5      $ 	#774FC8D4��6   $ 0.797Q�6-   � 5      $ 	#774FC8D4��6   $ 0.291Q�6-   5      �� 5   
   $ 0.5,1&�66   $ 0.5,0&�6G   .��5      � $ 	#FF00E8E8��6   $ 0Q�6-   � 5      $ 	#FF2FB7E3��6   $ 0.523Q�6-   � 5      $ 	#FF00E8E8��6   $ 1Q�6-   5      5      .��5      � $	��308��6   $
��1344��6    � ColumnDefinitions 5   
   �� 5      $	��59* �6   �� 5      $	��56* �6   �� 5      $
��208* �6   �� 5      $
��145* �6   �� 5      $	��92* �6   �� 5      $	��80* �6   �� 5      $	��86* �6   �� 5      $	��71* �6   �� 5      $
��111* �6   �� 5       $	��67* �6   �� 5!      $	��78* �6   �� 5"      $
��101* �6   �� 5#      $	��79* �6   �� 5$      $	��85* �6   �� 5%      $	��83* �6    � RowDefinitions 5'   
   �� 5(      $	��53* �6   �� 5)      $	��32* �6   �� 5*      $
��113* �6   �� 5+      $
��161* �6   .+�5-   
   �� -   5        $��border2��5-   [   
�������$��1q�6&   $��2��6:   $��3��6J   ��6j   ��  �� Source# [  ._�5.      �� -   5        $��
textBlock1��5.   $   $��38��6   $6��0Наименование  мероприятия��66   $��WrapWithOverflow=�6W   $	��147��6w   $��Stretch=�6�   $��Center=�6�    �� TextAlignment$ Center=�6�   5/      �� 50   
   -   5        $��border3��50   J   
�������$��1q�6&   $��1��6:   $��3��6Y   ��6j   �� # [  ._�51      �� -   5        $��
textBlock3��51   $   $��93��6   $6��0Номер меро-приятия в плане��66   $��WrapWithOverflow=�6X   $ Center=�6x   $��Center=�6�   $��Center=�6�   52      �� 53   
   -   5        $��border4��53   :   
�������$��1q�6&   $��3��6I   ��6Z   �� # [  ._�54      �� -   5        $��
textBlock2��54   %   $	��138��6   $V��PКод  основных направлений энергосбережения��67   $ Center=�6i   $��Center=�6�   $��Center=�6�   $��49��6�   $
��Wrap=�6�   55      �� 56   
   -   5        $��border5��56   [   
�������$��1q�6&   $��3��6:   $��3��6J   ��6j   �� # [  ._�57      �� -   5        $��
textBlock4��57   $   $��42��6   $!��Дата внедрения��66   $ Center=�6L   $��WrapWithOverflow=�6c   $��Center=�6�   58      �� 59   
   -	   5        $��border6��59   5   $��1q�6   $��4��6%   $��2��6D   $��2��6�   ��6X   �� # [  ._�5:      �� -
   5        $��
textBlock5��5:   $   $��25��6   $#��Объем внедрения��66   $��Center=�6M   $��WrapWithOverflow=�6h   $ Center=�6�   5;      �� 5<   
   -   5        $��border7��5<   5   $��1q�6   $��4��6%   $��2��6D   ��6Q   �� # [  ._�5=      �� -   5        $��
textBlock6��5=   $   $��37��6   $'��!Единица измерения��66   $��Center=�6O   $ Center=�6j   $��WrapWithOverflow=�6�   5>      �� 5?   
   -   5        $��border8��5?   %   $��1q�6   $��6��64   $��3��6D   ��6U   �� # [  ._�5@      �� -   5        $��
textBlock8��5@   $   $��53��6   $4��.Экономия ТЭР, т. усл. топл.��66   $��Center=�6Y   $��WrapWithOverflow=�6t   $ Center=�6�   5A      �� 5B   
   -   5        $��border1��5B   k   
�������$��1q�6&   $��8��6:   $��7��6J   $��1��6^   ��6z   �� # [  ._�5C      �� -   5        $��textBlock11��5C   $   $��23��6   $=��7из них за счет средств бюджета��67   $��Center=�6]   $��Center=�6x   $��WrapWithOverflow=�6�   $ Center=�6�   5D      �� 5E   
   -   5        $��border10��5E   ^   
�������$��1q�6&   $��7��6:   $��8��6J   $��1,0,0,0q�6�   ��6n   �� # [  ._�5F      �� -   5        $��
textBlock9��5F   $   $��20��6   $����Затраты на внедрение мероприятия  и источники финансирования, млн. руб.��66   $��Center=�6�   $��WrapWithOverflow=�6�   $ Center=�6�   5G      �� 5H   
   -   5        $��border11��5H   W   
�������$��1q�6&   $��7��6:   $��1��6J   $��2��6g   ��6x   �� # [  ._�5I      �� -   5        $��textBlock10��5I   $   $��23��6   $��
Всего��67   $��Center=�6D   $ Center=�6_   $��Center=�6v   $��WrapWithOverflow=�6�   5J      �� 5K   
   -   5        $��border12��5K   W   
�������$��1q�6&   $��8��6:   $��2��6J   ��6g   �� # [  ._�5L      �� -   5        $��textBlock12��5L   $   $��93��6   $r��lресп. бюджета на финансирование программ энергосбережения��67   $��Center=�6x   $ Center=�6�   $��Stretch=�6�   $
��Wrap=�6�   $��84��6�   5M      �� 5N   
   -   5        $��border13��5N   W   
�������$��1q�6&   $��9��6:   $��2��6J   ��6g   �� # [  ._�5O      �� -   5        $��textBlock13��5O   $   $��57��6   $/��)инвестиционных фондов��67   $��Center=�6T   $
��Wrap=�6o   $ Center=�6�   $��Center=�6�   5P      �� 5Q   
   -   5        $��border14��5Q   X   
�������$��1q�6&   $��10��6:   $��2��6K   ��6h   �� # [  ._�5R      �� -   5        $��textBlock14��5R   $   $��59��6   $5��/республиканского бюджета��67   $��Center=�6W   $ Center=�6r   $
��Wrap=�6�   5S      �� 5T   
   -   5        $��border15��5T   X   
�������$��1q�6&   $��11��6:   $��2��6K   ��6h   �� # [  ._�5U      �� -   5        $��textBlock15��5U   $   $��34��6   $%��местного бюджета��67   $��Center=�6O   $
��Wrap=�6j   $ Center=�6~   5V      �� 5W   
   -   5        $��border16��5W   X   
�������$��1q�6&   $��12��6:   $��2��6K   ��6h   �� # [  ._�5X      �� -   5        $��textBlock16��5X   $   $��25��6   $��организации��67   $��Center=�6J   5Y      �� 5Z   
   -   5        $��border17��5Z   X   
�������$��1q�6&   $��13��6:   $��2��6K   ��6h   �� # [  ._�5[      �� -    5        $��textBlock17��5[   $   $��23��6   $��кредита��67   $��Center=�6F   $ Center=�6a   $
��Wrap=�6x   5\      �� 5]   
   -!   5        $��border18��5]   X   
�������$��1q�6&   $��14��6:   $��2��6K   ��6h   �� # [  ._�5^      �� -"   5        $��textBlock18��5^   $   $��48��6   $'��!Других источников��67   $��Center=�6P   $
��Wrap=�6k   $ Center=�6   $��53��6�   5_      �� 5`   
   -#   5        $��gridSplitter1��5`   A   $��5��6   $��2��6'   $	��113��64   $	��Top=�6V   $��0,0,1,0q�6�   ��6n   �� # [  6�   �� 5a   
   -$   5        $��gridSplitter2��5a   4   $��5��6   $��1��6'   $��0,30,1,0q�6�   $��2��6�   ��6I   �� # [  6�   �� 5b   
   -%   5        $��border9��5b   B   $��1q�6   $��5��6%   $��2��65   $��0,0,1,0q�6Q   ��6b   �� # [  ._�5c      �� -&   5        $��
textBlock7��5c      $��Количество��6*   $��46��6<   $
��Wrap=�6G   $ Center=�6[   	 �� IsHyphenationEnabled	 .  $��36��6�   5d      5e      5f      