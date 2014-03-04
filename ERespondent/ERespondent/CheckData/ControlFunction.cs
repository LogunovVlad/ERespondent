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
        public ControlFunction()
        { }

        List<string> listError = new List<string>();
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
                for (int i = 0; i < _rowCount - 1; i++)
                {
                    double allSumm = Convert.ToDouble(grid.Rows[i].Cells[7].Value);
                    double summItem = Convert.ToDouble(grid.Rows[i].Cells[8].Value) + Convert.ToDouble(grid.Rows[i].Cells[9].Value) +
                       Convert.ToDouble(grid.Rows[i].Cells[10].Value) + Convert.ToDouble(grid.Rows[i].Cells[11].Value) +
                       Convert.ToDouble(grid.Rows[i].Cells[12].Value) + Convert.ToDouble(grid.Rows[i].Cells[13].Value) + Convert.ToDouble(grid.Rows[i].Cells[14].Value);
                    if (allSumm != summItem)
                    {
                        listError.Add("\n<<" + grid.Tag.ToString() + ">>");
                        string stError = "Ошибка: В строке " + (i + 1) + " сумма столбцов {4-10} не равна значению в столбце <3>! Проверьте данные!";
                        listError.Add(stError);
                        grid.Rows[i].Cells[7].Style.BackColor = Color.LightSteelBlue;
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
            else
            {
                listError.Add(grid.Tag.ToString() + " не заполнен!\n");
            }
        }

        /// <summary>
        /// Проверяет сумму по столбцам
        /// </summary>
        /// <param name="grid"></param>
        private void AllBudgetColumn(DataGridView grid)
        {
            int _rowCount = grid.RowCount;
            int _columnCount = grid.ColumnCount;

            double _summColumn = 0;
            for (int j = 6; j < _columnCount; j++)
            {
                _summColumn = 0;
                //_rowCount-2 -> потому что последняя строка - "Итого"
                for (int i = 0; i < _rowCount - 1; i++)
                {
                    _summColumn += Convert.ToDouble(grid.Rows[i].Cells[j].Value);
                }
                if (_summColumn != Convert.ToDouble(grid.Rows[_rowCount - 1].Cells[j].Value))
                {
                    listError.Add("Ошибка: В столбце <" + grid.Columns[j].HeaderText + "> сумма пунктов не равна значению <Итого>! Проверьте данные!");
                    grid.Rows[_rowCount - 1].Cells[j].Style.BackColor = Color.Red;
                }
                else
                {
                    grid.Rows[_rowCount - 1].Cells[j].Style.BackColor = Color.White;
                }
            }

        }



        /// <summary>
        /// Выполняет проверку Итого по горизонтали и вертикали
        /// </summary>
        /// <param name="grid"></param>
        public void CheckTable(DataGridView grid)
        {
            AllBudgetRow(grid);
            AllBudgetColumn(grid);
        }

        /// <summary>
        /// Контрольные функции по разделу 1
        /// </summary>
        /// <param name="grid1">1. По плану мероприятий отчетного года</param>
        /// <param name="grid2">2. Дополнительные мероприятия</param>
        /// <param name="grid3">3. По мероприятиям предшествующего года внедрения</param>
        public void CheckTotalForSection1(DataGridView grid1, DataGridView grid2, DataGridView grid3)
        {
            int _rowCount = grid3.RowCount;
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
            #region
            for (int i = 6; i < 15; i++)
            {
                _summ = Convert.ToDouble(grid3[i, grid3.RowCount - 1].Value);
                double total1 = Convert.ToDouble(grid1[i, grid1.RowCount - 1].Value);
                double total2 = Convert.ToDouble(grid2[i, grid2.RowCount - 1].Value);              

                if (_summ != total1 + total2)
                {
                    listError.Add("Ошибка: Сумма данных <Всего по разделу 1> по столбцу "+grid3.Columns[i].HeaderText+" не равна сумме данных строк <Итого> по каждому из подразделов.");
                    grid3[i, grid3.RowCount - 1].Style.BackColor = Color.Red;
                }
                else
                {
                    grid3[i, grid3.RowCount - 1].Style.BackColor = Color.White;
                }
            }
            #endregion
        }

        public void ShowListError()
        {
            CheckProtocol formProtocol = new CheckProtocol(listError);
            formProtocol.Show();
        }
    }
}
