using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;

namespace ERespondent.CheckData
{
    class ControlFunction
    {
        CheckProtocol formProtocol;
        private List<string> listError;// = new List<string>();

        public ControlFunction()
        {
            formProtocol = new CheckProtocol();
            listError = new List<string>();
        }

        /// <summary>
        /// Функция контроля суммы строки по бюджету
        /// </summary>
        /// <param name="grid"></param>
        /// item = 8(для первого раздела)
        /// item = 10(для второго раздела)
        private void AllBudgetRow(DataGridView grid, int item)
        {
            //grid.CurrentCell.Selected = false;
            grid.EndEdit();
            int _rowCount = grid.RowCount;
            int _columnCount = grid.ColumnCount;
            if (_rowCount > 1)
            {
                for (int i = 0; i < _rowCount; i++)
                {
                    double allSumm = Convert.ToDouble(grid.Rows[i].Cells[item - 1].Value);
                    double summItem = 0;
                    for (int iItem = item; iItem < _columnCount; iItem++)
                    {
                        summItem += Convert.ToDouble(grid.Rows[i].Cells[iItem].Value);
                    }
                    if (allSumm != summItem)
                    {
                        string stError = null;
                        switch (grid.Tag.ToString())
                        {
                            case "T1":
                                stError = "1. По плану мероприятий отчетного года";
                                break;
                            case "T2":
                                stError = "2. Дополнительные мероприятия";
                                break;
                            case "T3":
                                stError = "3. По мероприятиям предшествующего года внедрения";
                                break;
                        }
                        listError.Add("\n<<" + stError + ">>");

                        if (i == _rowCount - 1)
                        {
                            stError = "Ошибка: В строке " + (i + 1) +
                                " (ИТОГО) сумма столбцов {4-10} не равна значению в столбце <3>! Проверьте данные!";
                        }
                        else
                        {
                            stError = "Ошибка: В строке " + (i + 1) + " (код основных направлений: "
                               + grid[0, i].Value + ") сумма столбцов {4-10} не равна значению в столбце <3>! Проверьте данные!";
                        }
                        listError.Add(stError);
                        grid.Rows[i].Cells[item - 1].Style.BackColor = Color.Red;
                        for (int indexRow = item; indexRow < _columnCount; indexRow++)
                        {
                            grid.Rows[i].Cells[indexRow].Style.BackColor = Color.Yellow;
                        }
                    }
                    else
                    {
                        for (int indexRow = item; indexRow < _columnCount; indexRow++)
                        {
                            grid.Rows[i].Cells[indexRow - 1].Style.BackColor = Color.White;
                        }
                    }
                }
            }
            //else
            //{
            //    listError.Add(grid.Tag.ToString() + " не заполнен!\n");
            //}
        }

        /// <summary>
        /// Контрольные функции по разделу 1
        /// </summary>
        /// <param name="grid1">1. По плану мероприятий отчетного года</param>
        /// <param name="grid2">2. Дополнительные мероприятия</param>
        /// <param name="grid3">3. По мероприятиям предшествующего года внедрения</param>
        /// <param name="columnTotal">Индекс колонки всего (6 - для первого раздела ("Экономия ТЭР"); 8 - для второго раздела ("Увеличение МВТ"))</param>
        private void CheckTotalForSection1(DataGridView grid1, DataGridView grid2, DataGridView grid3, int columnTotal)
        {
            int _rowCount = grid3.RowCount;
            //проверка запускается, если в таблице есть записи
            if (_rowCount > 2)
            {
                double _summ = 0;
                for (int i = 0; i < _rowCount - 2; i++)
                {
                    _summ += Convert.ToDouble(grid3[columnTotal, i].Value);
                }
                //итого по подраздеру 3 раздела 1
                if (Convert.ToDouble(grid3[columnTotal, _rowCount - 2].Value) != _summ)
                {
                    listError.Add("\n<<" + grid3.Tag.ToString() + ">>");
                    listError.Add("Ошибка: В строке <Итого> сумма не соответствует сумме по соответствующим столбцам!");
                    grid3[columnTotal, _rowCount - 2].Style.BackColor = Color.Red;
                }
                else
                {
                    grid3[columnTotal, _rowCount - 2].Style.BackColor = Color.White;
                }
                //end

                //итого по разделу 1 - Экономия ТЭР              
                _summ = Convert.ToDouble(grid3[columnTotal, grid3.RowCount - 1].Value);
                double total1 = Convert.ToDouble(grid1[columnTotal, grid1.RowCount - 1].Value);
                double total2 = Convert.ToDouble(grid2[columnTotal, grid2.RowCount - 1].Value);
                double total3 = Convert.ToDouble(grid3[columnTotal, grid3.RowCount - 2].Value);
                if (_summ != (total1 + total2 + total3))
                {
                    listError.Add("Ошибка: Сумма данных <Всего по разделу> по столбцу " + grid3.Columns[columnTotal].HeaderText + " не равна сумме данных строк <Итого> по каждому из подразделов.");
                    grid3[columnTotal, grid3.RowCount - 1].Style.BackColor = Color.Red;
                }
                else
                {
                    grid3[columnTotal, grid3.RowCount - 1].Style.BackColor = Color.White;
                }
            }
        }

        /// <summary>
        /// Контрольные функции по разделу
        /// </summary>
        /// <param name="grid1">Таблица 1 подраздела</param>
        /// <param name="grid2">Таблица 2 подраздела</param>
        /// <param name="grid3">Таблица 3 подраздела</param>
        /// <param name="section">Название раздела</param>
        /// <param name="columnTotal">Номер столбца ВСЕГО. Для 1 раздела - 8; для второго - 10</param>
        public void CheckSection(DataGridView grid1, DataGridView grid2, DataGridView grid3, string section, int columnTotal)
        {
            AllBudgetRow(grid1, columnTotal);
            AllBudgetRow(grid2, columnTotal);
            //  убрал, т.к. сделал автоматическую подстановку для ВСЕГО ПО РАЗДЕЛУ.
            //  CheckTotalForSection1(grid1, grid2, grid3, columnTotal-2);
            CheckProtocol.ErrorForAllSection.Add(section, listError);
        }

        #region РАЗДЕЛ 3
        /// <summary>
        /// Проверяет равенство – фактическое количество мероприятий по экономии ТЭР и увеличению использования МВТ, 
        /// внедренных в отчетном периоде (количество мероприятий, включенных в позиции «1. По плану мероприятий отчетного
        /// года» и «2. Дополнительные мероприятия» разделов I и II)
        /// </summary>
        /// <param name="grid">Массив таблиц РАЗДЕЛА 1(подразделы 1 и 2) и РАЗДЕЛА 2(подразделы 1 и 2)</param>
        /// <param name="fact"></param>
        public void CheckSection3Table3Row1(DataGridView[] grid, DataGridView grid3)
        {
            int countAction = 0;
            int outI;

            for (int tableIndex = 0; tableIndex < grid.Count(); tableIndex++)
            {
                for (int i = 0; i < grid[tableIndex].RowCount - 1; i++)
                {
                    if (grid[i][0, i].Value != null)
                    {
                        countAction++;
                    }
                }
            }
            if (Convert.ToDouble(grid3[4, 1].Value) != countAction)
            {
                listError.Add("\nОшибка: Фактическое количество мероприятий по экономии" +
                " ТЭР и увеличению использования МВТ, внедренных в отчетном периоде не" +
                " соответствует количеству мероприятий в разделе 3 (таблица 3: <код строки <1>)");
                grid3[4, 1].Style.BackColor = Color.Red;
            }
            else
                grid3[4, 1].Style.BackColor = Color.White;
           
        }

        /// <summary>
        /// Проверяет фактическую величину экономии ТЭР, полученную путем суммирования итоговых значений в графе 2 
        /// по позициям «1. По плану мероприятий отчетного года» и «2. Дополнительные мероприятия» раздела I
        /// </summary>
        /// <param name="grid1"></param>
        /// <param name="grid2"></param>
        /// <param name="grid3"></param>
        public void CheckSection3Table3Row2(DataGridView grid1, DataGridView grid2, DataGridView grid3)
        {
            if ((Convert.ToDouble(grid1[6, grid1.RowCount - 1].Value) +
                Convert.ToDouble(grid2[6, grid2.RowCount - 1].Value)) != Convert.ToDouble(grid3[4, 2].Value))
            {
                listError.Add("\nОшибка: Фактическая величина экономии ТЭР (таблица 3: <код строки <2>>), не равна полученной сумме путем суммирования итоговых" +
                " значений в графе 2 по позициям «1. По плану мероприятий отчетного года» и «2. Дополнительные мероприятия» раздела I");
                grid1[6, grid1.RowCount - 1].Style.BackColor = Color.Red;
                grid2[6, grid2.RowCount - 1].Style.BackColor = Color.Red;
                grid3[4, 2].Style.BackColor = Color.Red;
            }
            else
            {
                grid1[6, grid1.RowCount - 1].Style.BackColor = Color.White;
                grid2[6, grid2.RowCount - 1].Style.BackColor = Color.White;
                grid3[4, 2].Style.BackColor = Color.White;
            }
        }

        /// <summary>
        /// Проверяет равенство – фактической величины увеличения использования МВТ, полученной путем суммирования итоговых 
        /// значений в графе 2 по позициям «1. По плану мероприятий отчетного года» и «2. Дополнительные мероприятия» раздела II
        /// </summary>
        /// <param name="grid1"></param>
        /// <param name="grid2"></param>
        /// <param name="grid3"></param>
        public void CheckSection3Table3Row3(DataGridView grid1, DataGridView grid2, DataGridView grid3)
        {
            if ((Convert.ToDouble(grid1[8, grid1.RowCount - 1].Value) +
                           Convert.ToDouble(grid2[8, grid2.RowCount - 1].Value)) != Convert.ToDouble(grid3[4, 3].Value))
            {
                listError.Add("\nОшибка: Фактическая величина увеличения использования МВТ <таблица 3: код строки <3>>, не равна полученной сумме "+
                "итоговых значений в графе 2 по позициям «1. По плану мероприятий отчетного года» и "+
                " «2. Дополнительные мероприятия» раздела II");
                grid1[8, grid1.RowCount - 1].Style.BackColor = Color.Red;
                grid2[8, grid2.RowCount - 1].Style.BackColor = Color.Red;
                grid3[4, 3].Style.BackColor = Color.Red;
            }
            else
            {
                grid1[8, grid1.RowCount - 1].Style.BackColor = Color.White;
                grid2[8, grid2.RowCount - 1].Style.BackColor = Color.White;
                grid3[4, 3].Style.BackColor = Color.White;
            }
        }

        /// <summary>
        /// Проверяет равенство - фактической величины затрат на внедрение мероприятий, полученной путем суммирования значений
        /// в графе 3 по строкам «Всего по разделу I» и «Всего по разделу II»;
        /// </summary>
        /// <param name="grid1"></param>
        /// <param name="grid2"></param>
        /// <param name="grid3"></param>
        public void CheckSection3Table3Row4(DataGridView grid1, DataGridView grid2, DataGridView grid3)
        {
            double allSection1 = Convert.ToDouble(grid1[7, grid1.RowCount - 1].Value);
            double allSection2 = Convert.ToDouble(grid2[9, grid2.RowCount - 1].Value);
            double factAll = Convert.ToDouble(grid3[4, 4].Value);
            if ((allSection1 + allSection2) != factAll)
            {
                listError.Add("Ошибка: Фактическая величина затрат на внедрение мероприятий, полученная путем" +
                    " суммирования значений в графе 3 по строкам «Всего по разделу I» и «Всего по разделу II» не" +
                    " равна сумме указанной в таблице 3 (таблица 3: <код строки <4>>);");
                grid1[7, grid1.RowCount - 1].Style.BackColor = Color.Red;
                grid2[9, grid2.RowCount - 1].Style.BackColor = Color.Red;
                grid3[4, 4].Style.BackColor = Color.Red;
            }
            else
            {
                grid1[7, grid1.RowCount - 1].Style.BackColor = Color.White;
                grid2[9, grid2.RowCount - 1].Style.BackColor = Color.White;
                grid3[4, 4].Style.BackColor = Color.White;
            }
        }

        #endregion
        /// <summary>
        /// Вывод ошибок
        /// </summary>
        public void ShowListError()
        {
            formProtocol.Show();
        }

    }
}

