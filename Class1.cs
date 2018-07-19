using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.IO;
using System.Text;

namespace MatrixCOM
{
    [Guid("425570DC-6672-466A-A50C-C581E4AD9D7A")]
    public interface MCOM_Interface
    {
        [DispId(1)]
        double DeterminantSequential(double[,] matrix, int row);
        [DispId(2)]
        double DeterminantParallel(double[,] matrix, int row);
        [DispId(3)]
        double DeterminantSequentialString(string matrixS, int size);
        [DispId(4)]
        double DeterminantParallelString(string matrixS, int size);
    }

    [Guid("EEFF3518-17F2-4974-98C2-9E36BA29A124")
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface MCOM_Events
    {
    }

    [Guid("ADFBB477-1BF9-45A0-BCBD-89AA9AFD868F"),
    ClassInterface(ClassInterfaceType.None),
    ComSourceInterfaces(typeof(MCOM_Events))]
    public class MCOM_class : MCOM_Interface
    {
        public MCOM_class()
        {
        }

        #region Sequential_Loop
        public double DeterminantSequential(double[,] matrix, int row)
        {
            int size = (int)Math.Sqrt(matrix.Length); //Расчет размера матрицы
            double determinant = 0; //Обнуление переменной, хранящей детерминант

            if (size == 1) return matrix[0, 0]; //Если размер матрицы равен единице, то детерминант равен ее единственному члену

            int bestRow = betterRow(matrix); //Поиск строки в заданной матрице, содержащей наибольшее количество нулей
            double[,] excluded;

            for (int i = 0; i < size; i++)
            {
                if (matrix[bestRow, i] == 0) determinant += 0; //Если член строки равен нулю, то его произведение на минор также равно нулю
                else
                {
                    excluded = MatrixExclusion(matrix, bestRow, i); //Вычитание из заданной матрицы найденной строки с наибольшим количеством нулей 
                                                                    //и текущего столбца
                    determinant += matrix[bestRow, i] * Math.Pow(-1, (bestRow + i) + 2) * DeterminantSequential(excluded, 0); //Рекурсивный расчет детерминанта
                }
            }
            return determinant;
        }

        public double DeterminantSequentialString(string matrixS, int size)
        {
            double determinant = 0;
            double[,] matrix = MatrixRead(size, matrixS);
            
            if (size == 1) return matrix[0, 0];

            //
            int bestRow = betterRow(matrix);
            double[,] excluded;
            //

            for (int i = 0; i < size; i++)
            {
                if (matrix[bestRow, i] == 0) determinant += 0;
                else
                {
                    excluded = MatrixExclusion(matrix, bestRow, i);
                    determinant += matrix[bestRow, i] * Math.Pow(-1, (bestRow + i) + 2) * DeterminantSequential(excluded, 0);
                }
            }
            return determinant;
        }
        #endregion

        #region Parallel_Loop
        public double DeterminantParallel(double[,] matrix, int row)
        {
            int size = (int)Math.Sqrt(matrix.Length);
            double determinant = new double();
            double[] minor = new double[size];

            if (size == 1) return matrix[0, 0];

            //
            int bestRow = betterRow(matrix);
            double[,] excluded;
            //

            Parallel.For(0, size, i =>
            {
                if (matrix[bestRow, i] == 0) minor[i] = 0;
                else
                {
                    excluded = MatrixExclusion(matrix, bestRow, i);
                    minor[i] = DeterminantSequential(excluded, 0);
                }
            });
            for (int i = 0; i < size; i++)
            {
                determinant += matrix[bestRow, i] * Math.Pow(-1, (bestRow + i) + 2) * minor[i];
            }
            return determinant;
        }

        public double DeterminantParallelString(string matrixS, int size)
        {
            double determinant = new double();
            double[,] matrix = MatrixRead(size, matrixS);
            double[] minor = new double[size];

            if (size == 1) return matrix[0, 0];

            //
            int bestRow = betterRow(matrix);
            double[,] excluded;
            //

            Parallel.For(0, size, i =>
            {
                if (matrix[bestRow, i] == 0) minor[i] = 0;
                else
                {
                    excluded = MatrixExclusion(matrix, bestRow, i);
                    minor[i] = DeterminantSequential(excluded, 0);
                }
            });
            for (int i = 0; i < size; i++)
            {
                determinant += matrix[bestRow, i] * Math.Pow(-1, (bestRow + i) + 2) * minor[i];
            }
            return determinant;
        }
        #endregion

        private double[,] MatrixExclusion(double[,] matrix, int row, int column)
        {
            int size = (int)Math.Sqrt(matrix.Length);
            double[,] excluded = new double[size - 1, size - 1];
            int localRow = 0;
            int localColumn = 0;
            for (int i = 0; i < size; i++)
            {
                if (i != row)
                {
                    for (int j = 0; j < size; j++)
                    {
                        if (j != column)
                        {
                            excluded[localRow, localColumn] = matrix[i, j];
                            localColumn++;
                        }
                    }
                    localRow++;
                }
                localColumn = 0;

            }
            return excluded;
        }

        private int betterRow(double[,] matrix)
        {
            int row = 0;
            int counter = 0;
            int prevCounter;
            int size = (int)Math.Sqrt(matrix.Length);
            for (int i = 0; i < size; i++)
            {
                prevCounter = counter;
                counter = 0;
                for (int j = 0; j < size; j++)
                {
                    if (matrix[i, j] == 0) counter++;
                }
                if (counter >= prevCounter) row = i;
            }
            return row;
        }

        private double[,] MatrixRead(int size, string variables)
        {
            string[] str;
            double[,] matrix = new double[size, size];
            str = variables.Split(' ');
            for (int i = 0; i < size; i++)
            {
                for (int j = 0; j < size; j++)
                {
                    matrix[i, j] = double.Parse(str[i * size + j]);
                }
            }
            return matrix;
        }
    }
}
