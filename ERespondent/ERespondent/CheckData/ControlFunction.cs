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
        public ControlFunction()
        {
            formProtocol = new CheckProtocol();
        }

        private List<string> listError = new List<string>();
        /// <summary>
        /// Функция контроля суммы строки по бюджету
        /// </summary>
        /// <param name="grid"></param>
        private void AllBudgetRow(DataGridView grid)
        {
            //grid.CurrentCell.Selected = false;
            grid.EndEdit();
            int _rowCount = grid.RowCount;
            int _columnCount = grid.ColumnCount;
            if (_rowCount > 1)
            {
                for (int i = 0; i < _rowCount; i++)
                {
                    double allSumm = Convert.ToDouble(grid.Rows[i].Cells[7].Value);
                    double summItem = 0;
                    for (int iItem = 8; iItem < 15; iItem++)
                    {
                        summItem += Convert.ToDouble(grid.Rows[i].Cells[iItem].Value);
                    }
                    if (allSumm != summItem)
                    {
                        listError.Add("\n<<" + grid.Tag.ToString() + ">>");
                        string stError = null;
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
                        grid.Rows[i].Cells[7].Style.BackColor = Color.Red;
                        grid.Rows[i].Cells[8].Style.BackColor = Color.Yellow;
                        grid.Rows[i].Cells[9].Style.BackColor = Color.Yellow;
                        grid.Rows[i].Cells[10].Style.BackColor = Color.Yellow;
                        grid.Rows[i].Cells[11].Style.BackColor = Color.Yellow;
                        grid.Rows[i].Cells[12].Style.BackColor = Color.Yellow;
                        grid.Rows[i].Cells[13].Style.BackColor = Color.Yellow;
                        grid.Rows[i].Cells[14].Style.BackColor = Color.Yellow;
                    }
                    else
                    {
                        grid.Rows[i].Cells[7].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[8].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[9].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[10].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[11].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[12].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[13].Style.BackColor = Color.White;
                        grid.Rows[i].Cells[14].Style.BackColor = Color.White;
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
        private void CheckTotalForSection1(DataGridView grid1, DataGridView grid2, DataGridView grid3)
        {
            int _rowCount = grid3.RowCount;
            //проверка запускается, если в таблице есть записи
            if (_rowCount > 2)
            {
                double _summ = 0;
                for (int i = 0; i < _rowCount - 2; i++)
                {
                    _summ += Convert.ToDouble(grid3[6, i].Value);
                }
                //итого по подраздеру 3 раздела 1
                if (Convert.ToDouble(grid3[6, _rowCount - 2].Value) != _summ)
                {
                    listError.Add("\n<<" + grid3.Tag.ToString() + ">>");
                    listError.Add("Ошибка: В строке <Итого> сумма не соответствует сумме по соответствующим столбцам!");
                    grid3[6, _rowCount - 2].Style.BackColor = Color.Red;
                }
                else
                {
                    grid3[6, _rowCount - 2].Style.BackColor = Color.White;
                }
                //end

                //итого по разделу 1 - Экономия ТЭР              
                _summ = Convert.ToDouble(grid3[6, grid3.RowCount - 1].Value);
                double total1 = Convert.ToDouble(grid1[6, grid1.RowCount - 1].Value);
                double total2 = Convert.ToDouble(grid2[6, grid2.RowCount - 1].Value);
                double total3 = Convert.ToDouble(grid3[6, grid2.RowCount - 2].Value);
                if (_summ != (total1 + total2 + total3))
                {
                    listError.Add("Ошибка: Сумма данных <Всего по разделу 1> по столбцу " + grid3.Columns[6].HeaderText + " не равна сумме данных строк <Итого> по каждому из подразделов.");
                    grid3[6, grid3.RowCount - 1].Style.BackColor = Color.Red;
                }
                else
                {
                    grid3[6, grid3.RowCount - 1].Style.BackColor = Color.White;
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
        public void CheckSection(DataGridView grid1, DataGridView grid2, DataGridView grid3, string section)
        {
            AllBudgetRow(grid1);
            AllBudgetRow(grid2);
            CheckTotalForSection1(grid1, grid2, grid3);
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
