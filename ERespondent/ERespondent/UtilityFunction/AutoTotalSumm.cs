using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ERespondent.UtilityFunction
{
    class AutoTotalSumm
    {
        public static List<double> listTotalSummPoint;
        ///
        ///Методы вычисляющие итого по подразделу 1 и 2
        ///
        #region
        /// <summary>
        /// Метод вычисляющий сумму строки итого по подразделам
        /// </summary>
        /// <param name="grid">Текущая таблица</param>
        /// <param name="varSub">value "1" - если это таблица 1 или 2. Иначе value "2"</param>
        public static void TotalSumm(DataGridView grid)
        {
            listTotalSummPoint = new List<double>();
            double summTotal;

            for (int indexCol = 6; indexCol < grid.ColumnCount; indexCol++)
            {
                try
                {
                    summTotal = 0;
                    for (int indexRow = 0; indexRow < grid.RowCount - 1; indexRow++)
                    {
                        summTotal += Convert.ToDouble(grid[indexCol, indexRow].Value);
                    }
                    listTotalSummPoint.Add(summTotal);
                }
                catch (FormatException)
                {
                    MessageBox.Show("Ошибка формата!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    grid.CurrentCell.Value = null;
                    break;
                }
            }

        }

        /// <summary>
        /// Заполнение строки Итого подсчитанными значениями
        /// </summary>
        /// <param name="grid">Текущая таблица</param>
        /// <param name="varSub">>value "1" - если это таблица 1 или 2. Иначе value "2"</param>
        public static void FillTotalRow(DataGridView grid)
        {
            //пишем с 6-го столбца, т.к. начиная с него ведется расчет "Итого"
            int i = 6;
            foreach (double point in listTotalSummPoint)
            {
                if (point != 0)
                {
                    grid[i, grid.RowCount - 1].Value = point;
                }
                i++;
            }
        }
        #endregion

        /// <summary>
        /// Считает "Итого" в третьей таблице
        /// </summary>
        /// <param name="grid"></param>
        public static void TotalSummGrid3(DataGridView grid)
        {
            double summTotal = 0;
            for (int i = 0; i < grid.RowCount - 2; i++)
            {
                summTotal += Convert.ToDouble(grid[6, i].Value);
            }
            grid[6, grid.RowCount - 2].Value = summTotal;
        }

        /// <summary>
        /// Всего по разделу 1
        /// </summary>
        /// <param name="grid1"></param>
        /// <param name="grid2"></param>
        public static void TotalAll1Section(DataGridView grid1, DataGridView grid2, DataGridView grid3)
        {
            listTotalSummPoint = new List<double>();                 
            grid3.EndEdit();

            //ячейка итого по подразделу 3
            double s2 = Convert.ToDouble(grid3[6, grid3.RowCount - 2].Value);

            for (int i = 6; i < grid1.ColumnCount; i++)
            {
                double s=Convert.ToDouble(grid1[i, grid1.RowCount - 1].Value);
                double s1 = Convert.ToDouble(grid2[i, grid2.RowCount - 1].Value);
               
                listTotalSummPoint.Add(s+s1+s2);
                s2 = 0;
            }
        }

        /// <summary>
        /// Заполнение таблицы 3
        /// </summary>
        /// <param name="grid"></param>
        public static void FillGrid3(DataGridView grid)
        {
            int i = 6;            
            foreach (double point in listTotalSummPoint)
            {
                if (point != 0)
                {
                    grid[i, grid.RowCount - 1].Value = point;
                }
                i++;
            }
        }
            

    }
}
