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
        #region По разделу 1
        
        #region Методы вычисляющие итого по подразделам 1 и 2
        /// <summary>
        /// Метод вычисляющий сумму строки итого по пунктам
        /// </summary>
        /// <param name="grid">Текущая таблица</param>  
        /// <param name="startColumn">Столбец, с которого начинается расчет ИТОГО</param> 
        public static void TotalSumm(DataGridView grid, int startColumn)
        {
            listTotalSummPoint = new List<double>();
            double summTotal;

            for (int indexCol = startColumn; indexCol < grid.ColumnCount; indexCol++)
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
                    ///
                    /// поставить проверку на endEdit таблицы
                    /// 
                    MessageBox.Show("Ошибка формата: введите число!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    grid.CurrentCell.Value = null;
                    break;
                }
            }

        }

        /// <summary>
        /// Заполнение строки Итого подсчитанными значениями
        /// </summary>
        /// <param name="grid">Текущая таблица</param>     
        /// <param name="startColumn">Столбец, с которого начинаем заполнение</param> 
        public static void FillTotalRow(DataGridView grid, int startColumn)
        {                        
            foreach (double point in listTotalSummPoint)
            {
                if (point != 0)
                {
                    grid[startColumn, grid.RowCount - 1].Value = point;
                }
              /*  else
                {
                    grid[startColumn, grid.RowCount - 1].Value = 0;
                }*/
                startColumn++;
            }
        }
        #endregion

        /// <summary>
        /// Считает "Итого" в третьей таблице
        /// </summary>
        /// <param name="grid"></param>
        public static void TotalSummGrid3(DataGridView grid, int columnTotal)
        {
            double summTotal = 0;
            for (int i = 0; i < grid.RowCount - 2; i++)
            {
                summTotal += Convert.ToDouble(grid[columnTotal, i].Value);
            }
            grid[columnTotal, grid.RowCount - 2].Value = summTotal;
        }

        /// <summary>
        /// Всего по разделу 1
        /// </summary>
        /// <param name="grid1">таблица 1 (раздел1)</param>
        /// <param name="grid2">таблица 2 (раздел1)</param>
        /// <param name="grid3">таблица 3 (раздел1)</param>
        public static void TotalAll1Section(DataGridView grid1, DataGridView grid2, DataGridView grid3, int startColumn)
        {
            listTotalSummPoint = new List<double>();                 
            grid3.EndEdit();

            //ячейка итого по подразделу 3
            //double s2 = Convert.ToDouble(grid3[6, grid3.RowCount - 2].Value);
            double s2 = Convert.ToDouble(grid3[startColumn, grid3.RowCount - 2].Value);

            for (int i = startColumn; i < grid1.ColumnCount; i++)
            {
                double s = Convert.ToDouble(grid1[i, grid1.RowCount - 1].Value);
                double s1 = Convert.ToDouble(grid2[i, grid2.RowCount - 1].Value);
               
                listTotalSummPoint.Add(s+s1+s2);
                s2 = 0;
            }
        }

        /// <summary>
        /// Заполнение таблицы 3
        /// </summary>
        /// <param name="grid"></param>
        public static void FillGrid3(DataGridView grid, int columnFill)
        {                    
            foreach (double point in listTotalSummPoint)
            {
                if (point != 0)
                {
                    grid[columnFill, grid.RowCount - 1].Value = point;
                }
                columnFill++;
            }
        }
        #endregion      

    }
}
