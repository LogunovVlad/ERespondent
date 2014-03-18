using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

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
                    double allSumm = Convert.ToDouble(grid.Rows[i].Cells[item-1].Value);
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
                        grid.Rows[i].Cells[item-1].Style.BackColor = Color.Red;
                        for (int indexRow = item; indexRow < _columnCount; indexRow++)
                        {
                            grid.Rows[i].Cells[indexRow].Style.BackColor = Color.Yellow;
                        }                    
                    }
                    else
                    {
                        for (int indexRow = item; indexRow < _columnCount; indexRow++)
                        {
                            grid.Rows[i].Cells[indexRow-1].Style.BackColor = Color.White;
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

        /// <summary>
        /// Вывод ошибок
        /// </summary>
        public void ShowListError()
        {
            formProtocol.Show();
        }
    }
}
