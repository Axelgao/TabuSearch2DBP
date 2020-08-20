using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Drawing.Drawing2D;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace TabuSearch2DBP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "50";
        }

        // File location
        string fname = "D:\\Files\\Doc\\Waikato Uni\\556 Metaheuristic Algorithms\\TabuSearch2DBP\\TabuSearch2DBP\\Data\\556 Assignment Problem Instances.xlsx";

        // 2D rectangle drawing
        Pen border = new Pen(Color.Red, 1);
        SolidBrush fill = new SolidBrush(Color.Green);

        private void button1_Click(object sender, EventArgs e)
        {
            label1.Text = "Calculating M1a!";

            // Calculate the search time
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            #region Read Excel Data
            // Open excel file
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            System.Drawing.Rectangle[] rectangles = new System.Drawing.Rectangle[200];

            int regionHeightLimit = 0;
            int regionWidthLimit = 40;

            // Read rectangle data from excel file
            for (int i = 6; i <= 105; i++)
            {
                int number = 0;
                int width = 0;
                int height = 0;

                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 1)
                    {
                        number = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                    else if (j == 2)
                    {
                        width = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                    else if (j == 3)
                    {
                        height = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                }

                rectangles[number] = new System.Drawing.Rectangle(0, 0, width, height);

                if (width >= height)
                {
                    regionHeightLimit += width;
                }
                else
                {
                    regionHeightLimit += height;
                }
            }

            #endregion

            #region Create Initial Solution
            // Rectangle filling methodology: Bottom-up Left-justified
            // Create bool array to check if the space is available to arrange a rectangle
            bool[,] regionAvailable = InitialRegion(regionWidthLimit, regionHeightLimit);

            
            int rectangleCount = 100;

            // Create filling order list including 3 parts: 
            // 1.Total Height figure
            // 2.Orders for the rectangles
            // 3.Non-rotation or Rotation (1 or 2)
            // like: [Total Height,    1,3,4,2,5....99,100,   1,2,1,2,1,1,2,2,2......]
            // RANDOMLY generate initial filling order solution with temporary Total Height 0 and all non-rotation 1
            int[] rectangleFillingHeightOrder = InitialRandomOrder(rectangleCount);

            // Filling the rectangles into the region to calculate the Total Height
            rectangleFillingHeightOrder[0] = CalculateTotalHeight(rectangles, rectangleFillingHeightOrder, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);

            // Count solution
            int solutionCount = 1;

            #endregion

            #region Tabu Search
            // Tabu search termination criteria low height
            int terminationHeight = 63; // Calculation: Rectangles Area / Rigion Width

            // Create Tabu list
            int tabuLength = Convert.ToInt32(textBox1.Text); // Tabu length
            int[][] tabu = new int[tabuLength][];
            for (int t = 0; t < tabuLength; t++)
            {
                tabu[t] = new int[rectangleCount * 2 + 1];
            }

            // Initial the tabu list with the start solution
            Array.Copy(rectangleFillingHeightOrder, tabu[0], rectangleCount * 2 + 1);

            // Count how many tabus now
            int tabuCount = 1;

            // Current solution
            int[] currentSolution = new int[rectangleCount * 2 + 1];
            Array.Copy(rectangleFillingHeightOrder, currentSolution, rectangleCount * 2 + 1);

            // Best solution
            int[] bestSolution = new int[rectangleCount * 2 + 1];
            Array.Copy(rectangleFillingHeightOrder, bestSolution, rectangleCount * 2 + 1);

            // Set the max moves
            int maxMoves = 10000;

            // Start the moving loop
            for (int m = 1; m <= maxMoves; m++)
            {
                // Best neighbor solution
                int[] bestNeighbor = new int[rectangleCount * 2 + 1];
                bool bestNeighborIsEmpty = true;

                // Randomly choose one rectangle to swap the order 
                Random rnd = new Random();
                int tmp = rnd.Next(1, rectangleCount + 1);

                // Loop for each neighbor
                for (int i = 1; i <= rectangleCount; i++)
                {
                    if (i == tmp)
                    {
                        continue;
                    }

                    int[] neighbor = new int[rectangleCount * 2 + 1];
                    Array.Copy(currentSolution, neighbor, rectangleCount * 2 + 1);

                    int swap = neighbor[i];
                    neighbor[i] = neighbor[tmp];
                    neighbor[tmp] = swap;

                    Random rotation = new Random();
                    int rotationNumber = rotation.Next(1, 3);

                    neighbor[i + rectangleCount] = rotationNumber;

                    neighbor[0] = CalculateTotalHeight(rectangles, neighbor, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);
                    solutionCount++;


                    // Check if the solution is the best solution found ever
                    if (neighbor[0] < bestSolution[0])
                    {
                        Array.Copy(neighbor, bestSolution, rectangleCount * 2 + 1);

                        // Check if it's the Termination Height value
                        if (bestSolution[0] == terminationHeight)
                        {
                            break;
                        }
                    }

                    // Check if the neighbor solution is already in tabu list
                    bool inTabu = false;

                    for (int t = 0; t < tabuCount; t++)
                    {
                        bool checkNextTabu = false;

                        for (int c = 1; c <= rectangleCount * 2; c++)
                        {
                            if (tabu[t][c] != neighbor[c])
                            {
                                // this neighbor is not this item
                                checkNextTabu = true;
                                break;
                            }
                        }

                        if (checkNextTabu)
                        {
                            continue;
                        }

                        // The order and rotation of tabu[t] is the same as neighbor's 
                        // This neighbor solution is already in tabu list
                        inTabu = true;
                    }

                    // If not in tabu list, record the neighbor solution
                    if (!inTabu)
                    {
                        if (bestNeighborIsEmpty)
                        {
                            // First admitted neighbor
                            Array.Copy(neighbor, bestNeighbor, rectangleCount * 2 + 1);
                            bestNeighborIsEmpty = false;
                        }
                        else
                        {
                            if (bestNeighbor[0] > neighbor[0])
                            {
                                Array.Copy(neighbor, bestNeighbor, rectangleCount * 2 + 1);
                            }
                        }
                    }
                }

                // There is no neighbor better than current solution
                if (currentSolution[0] <= bestNeighbor[0])
                {
                    Array.Copy(currentSolution, tabu[tabuCount], rectangleCount * 2 + 1);

                    if (tabuCount < tabuLength - 1)
                    {
                        tabuCount++;
                    }
                    else
                    {
                        tabuCount++;
                        break;
                    }
                }

                // Move to the best neighbor
                Array.Copy(bestNeighbor, currentSolution, rectangleCount * 2 + 1);
            }

            // Rerun bestSolution to generate the rectangles' positions saved in rectangle list
            bestSolution[0] = CalculateTotalHeight(rectangles, bestSolution, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);

            #endregion

            #region Draw Best Solution

            // Draw the best solution on the form
            for (int r = 1; r <= rectangleCount; r++)
            {
                int recNumber = bestSolution[r];

                System.Drawing.Rectangle shape = rectangles[recNumber];

                int x = shape.X;
                int y = shape.Y;
                int w = shape.Width;
                int h = shape.Height;

                System.Drawing.Rectangle shapeDraw;

                if (bestSolution[r + rectangleCount] == 1)
                {
                    // non-rotation
                    shapeDraw = new System.Drawing.Rectangle(x, y, w, h);
                }
                else
                {
                    // rotation
                    shapeDraw = new System.Drawing.Rectangle(x, y, h, w);
                }

                Graphics g2 = CreateGraphics();
                g2.FillRectangle(fill, shapeDraw);

                Graphics g1 = CreateGraphics();
                g1.DrawRectangle(border, shapeDraw);
            }

            #endregion

            #region Set Screen Info

            label1.Text = "M1a Calculation Done!";
            label2.Text = "Solution Compared:" + solutionCount.ToString();
            label3.Text = "Tabu:" + tabuCount.ToString();
            label4.Text = "Best Height:" + bestSolution[0].ToString();
            label7.Text = "Winforms' top left corner is (0,0). Will change picture orientation in the word report.";

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            label8.Text = "Search Time: " + elapsedTime;
            #endregion
        }

        private void button2_Click(object sender, EventArgs e)
        {
            label1.Text = "Calculating M2c!";

            // Calculate the search time
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            #region Read Excel Data
            // Open excel file
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            System.Drawing.Rectangle[] rectangles = new System.Drawing.Rectangle[200];

            int regionHeightLimit = 0;
            int regionWidthLimit = 100;

            // Read rectangle data from excel file
            for (int i = 6; i <= 105; i++)
            {
                int number = 0;
                int width = 0;
                int height = 0;

                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 5)
                    {
                        number = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                    else if (j == 6)
                    {
                        width = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                    else if (j == 7)
                    {
                        height = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                }

                rectangles[number] = new System.Drawing.Rectangle(0, 0, width, height);

                if (width >= height)
                {
                    regionHeightLimit += width;
                }
                else
                {
                    regionHeightLimit += height;
                }
            }


            #endregion

            #region Create Initial Solution

            // Rectangle filling methodology: Bottom-up Left-justified
            // Create bool array to check if the space is available to arrange a rectangle
            bool[,] regionAvailable = InitialRegion(regionWidthLimit, regionHeightLimit);

            int rectangleCount = 100;

            // Create filling order list including 3 parts: 
            // 1.Total Height figure
            // 2.Orders for the rectangles
            // 3.Non-rotation or Rotation (1 or 2)
            // like: [Total Height,    1,3,4,2,5....99,100,   1,2,1,2,1,1,2,2,2......]
            // RANDOMLY generate initial filling order solution with temporary Total Height 0 and all non-rotation 1
            int[] rectangleFillingHeightOrder = InitialRandomOrder(rectangleCount);

            // Filling the rectangles into the region to calculate the Total Height
            rectangleFillingHeightOrder[0] = CalculateTotalHeight(rectangles, rectangleFillingHeightOrder, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);

            // Count solution
            int solutionCount = 1;

            #endregion

            #region Tabu Search
            // Tabu search termination criteria low height
            int terminationHeight = 252; // Calculation: Rectangles Area / Rigion Width


            // Create Tabu list
            int tabuLength = Convert.ToInt32(textBox1.Text); // Tabu length
            int[][] tabu = new int[tabuLength][];
            for (int t = 0; t < tabuLength; t++)
            {
                tabu[t] = new int[rectangleCount *2 + 1];
            }

            // Initial the tabu list with the start solution
            Array.Copy(rectangleFillingHeightOrder, tabu[0], rectangleCount * 2 + 1);

            // Count how many tabus now
            int tabuCount = 1;

            // Current solution
            int[] currentSolution = new int[rectangleCount * 2 + 1];
            Array.Copy(rectangleFillingHeightOrder, currentSolution, rectangleCount * 2 + 1);

            // Best solution
            int[] bestSolution = new int[rectangleCount * 2 + 1];
            Array.Copy(rectangleFillingHeightOrder, bestSolution, rectangleCount * 2 + 1);

            // Set the max moves
            int maxMoves = 10000;

            // Start the moving loop
            for (int m = 1; m <= maxMoves; m++)
            {
                // Best neighbor solution
                int[] bestNeighbor = new int[rectangleCount * 2 + 1];
                bool bestNeighborIsEmpty = true;

                // Randomly choose one rectangle to swap the order 
                Random rnd = new Random();
                int tmp = rnd.Next(1,rectangleCount+1);

                // Loop for each neighbor
                for (int i = 1; i <= rectangleCount; i++)
                {
                    if (i == tmp)
                    {
                        continue;
                    }

                    int[] neighbor = new int[rectangleCount * 2 + 1];
                    Array.Copy(currentSolution, neighbor, rectangleCount * 2 + 1);

                    int swap = neighbor[i];
                    neighbor[i] = neighbor[tmp];
                    neighbor[tmp] = swap;

                    Random rotation = new Random();
                    int rotationNumber = rotation.Next(1, 3);

                    neighbor[i + rectangleCount] = rotationNumber;

                    neighbor[0] = CalculateTotalHeight(rectangles, neighbor, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);
                    solutionCount++;
                    

                    // Check if the solution is the best solution found ever
                    if (neighbor[0] < bestSolution[0])
                    {
                        Array.Copy(neighbor, bestSolution, rectangleCount * 2 + 1);

                        // Check if it's the Termination Height value
                        if (bestSolution[0] == terminationHeight)
                        {
                            break;
                        }
                    }

                    // Check if the neighbor solution is already in tabu list
                    bool inTabu = false;

                    for (int t = 0; t < tabuCount; t++ )
                    {
                        bool checkNextTabu = false;

                        for (int c = 1; c <= rectangleCount * 2; c++)
                        {
                            if (tabu[t][c] != neighbor[c])
                            {
                                // this neighbor is not this item
                                checkNextTabu = true;
                                break;
                            }
                        }

                        if (checkNextTabu)
                        {
                            continue;
                        }

                        // The order and rotation of tabu[t] is the same as neighbor's 
                        // This neighbor solution is already in tabu list
                        inTabu = true;  
                    }

                    // If not in tabu list, record the neighbor solution
                    if (!inTabu)
                    {
                        if (bestNeighborIsEmpty)
                        {
                            // First admitted neighbor
                            Array.Copy(neighbor, bestNeighbor, rectangleCount * 2 + 1);
                            bestNeighborIsEmpty = false;
                        }
                        else
                        {
                            if (bestNeighbor[0] > neighbor[0])
                            {
                                Array.Copy(neighbor, bestNeighbor, rectangleCount * 2 + 1);
                            }
                        }
                    }
                }

                // There is no neighbor better than current solution
                if (currentSolution[0] <= bestNeighbor[0])
                {
                    Array.Copy(currentSolution, tabu[tabuCount], rectangleCount * 2 + 1);

                    if (tabuCount < tabuLength - 1)
                    {
                        tabuCount++;
                    }
                    else
                    {
                        tabuCount++;
                        break;
                    }
                }

                // Move to the best neighbor
                Array.Copy(bestNeighbor, currentSolution, rectangleCount * 2 + 1);
            }

            // Rerun bestSolution to generate the rectangles' positions saved in rectangle list
            bestSolution[0] = CalculateTotalHeight(rectangles, bestSolution, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);

            #endregion

            #region Draw Best Solution

            // Draw the best solution on the form
            for (int r = 1; r <= rectangleCount; r++)
            {
                int recNumber = bestSolution[r];

                System.Drawing.Rectangle shape = rectangles[recNumber];

                int x = shape.X;
                int y = shape.Y;
                int w = shape.Width;
                int h = shape.Height;

                System.Drawing.Rectangle shapeDraw;

                if (bestSolution[r + rectangleCount] == 1)
                {
                    // non-rotation
                    shapeDraw = new System.Drawing.Rectangle(x, y, w, h);
                }
                else
                {
                    // rotation
                    shapeDraw = new System.Drawing.Rectangle(x, y, h, w);
                }

                Graphics g2 = CreateGraphics();
                g2.FillRectangle(fill, shapeDraw);

                Graphics g1 = CreateGraphics();
                g1.DrawRectangle(border, shapeDraw);
            }

            #endregion

            #region Set Screen Info
            label1.Text = "M2c Calculation Done!";
            label2.Text = "Solution Compared:" + solutionCount.ToString();
            label3.Text = "Tabu:" + tabuCount.ToString();
            label4.Text = "Best Height:" + bestSolution[0].ToString();
            label7.Text = "Winforms' top left corner is (0,0). Will change picture orientation in the word report.";

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            label8.Text = "Search Time: " + elapsedTime;
            #endregion
        }



        private void button3_Click(object sender, EventArgs e)
        {
            label1.Text = "Calculating M3d!";

            // Calculate the search time
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            #region Read Excel File

            // Open excel file
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            System.Drawing.Rectangle[] rectangles = new System.Drawing.Rectangle[200];

            int regionHeightLimit = 0;
            int regionWidthLimit = 100;

            // Read rectangle data from excel file
            for (int i = 6; i <= 155; i++)
            {
                int number = 0;
                int width = 0;
                int height = 0;

                for (int j = 1; j <= colCount; j++)
                {
                    if (j == 9)
                    {
                        number = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                    else if (j == 10)
                    {
                        width = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                    else if (j == 11)
                    {
                        height = Convert.ToInt32(xlRange.Cells[i, j].Value2);
                    }
                }

                rectangles[number] = new System.Drawing.Rectangle(0, 0, width, height);

                if (width >= height)
                {
                    regionHeightLimit += width;
                }
                else
                {
                    regionHeightLimit += height;
                }
            }

            #endregion

            #region Create Initial Solution

            // Rectangle filling methodology: Bottom-up Left-justified
            // Create bool array to check if the space is available to arrange a rectangle
            bool[,] regionAvailable = InitialRegion(regionWidthLimit, regionHeightLimit);

            int rectangleCount = 150;

            // Create filling order list including 3 parts: 
            // 1.Total Height figure
            // 2.Orders for the rectangles
            // 3.Non-rotation or Rotation (1 or 2)
            // like: [Total Height,    1,3,4,2,5....149,150,   1,2,1,2,1,1,2,2,2......]
            // RANDOMLY generate initial filling order solution with temporary Total Height 0 and all non-rotation 1
            int[] rectangleFillingHeightOrder = InitialRandomOrder(rectangleCount);

            // Filling the rectangles into the region to calculate the Total Height
            rectangleFillingHeightOrder[0] = CalculateTotalHeight(rectangles, rectangleFillingHeightOrder, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);

            // Count solution
            int solutionCount = 1;

            #endregion

            #region Tabu Search
            // Tabu search termination criteria low height
            int terminationHeight = 400; // Calculation: Rectangles Area / Rigion Width


            // Create Tabu list
            int tabuLength = Convert.ToInt32(textBox1.Text); // Tabu length
            int[][] tabu = new int[tabuLength][];
            for (int t = 0; t < tabuLength; t++)
            {
                tabu[t] = new int[rectangleCount * 2 + 1];
            }

            // Initial the tabu list with the start solution
            Array.Copy(rectangleFillingHeightOrder, tabu[0], rectangleCount * 2 + 1);

            // Count how many tabus now
            int tabuCount = 1;

            // Current solution
            int[] currentSolution = new int[rectangleCount * 2 + 1];
            Array.Copy(rectangleFillingHeightOrder, currentSolution, rectangleCount * 2 + 1);

            // Best solution
            int[] bestSolution = new int[rectangleCount * 2 + 1];
            Array.Copy(rectangleFillingHeightOrder, bestSolution, rectangleCount * 2 + 1);

            // Set the max moves
            int maxMoves = 10000;

            // Start the moving loop
            for (int m = 1; m <= maxMoves; m++)
            {
                // Best neighbor solution
                int[] bestNeighbor = new int[rectangleCount * 2 + 1];
                bool bestNeighborIsEmpty = true;

                // Randomly choose one rectangle to swap the order 
                Random rnd = new Random();
                int tmp = rnd.Next(1, rectangleCount + 1);

                // Loop for each neighbor
                for (int i = 1; i <= rectangleCount; i++)
                {
                    if (i == tmp)
                    {
                        continue;
                    }

                    int[] neighbor = new int[rectangleCount * 2 + 1];
                    Array.Copy(currentSolution, neighbor, rectangleCount * 2 + 1);

                    int swap = neighbor[i];
                    neighbor[i] = neighbor[tmp];
                    neighbor[tmp] = swap;

                    Random rotation = new Random();
                    int rotationNumber = rotation.Next(1, 3);

                    neighbor[i + rectangleCount] = rotationNumber;

                    neighbor[0] = CalculateTotalHeight(rectangles, neighbor, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);
                    solutionCount++;


                    // Check if the solution is the best solution found ever
                    if (neighbor[0] < bestSolution[0])
                    {
                        Array.Copy(neighbor, bestSolution, rectangleCount * 2 + 1);

                        // Check if it's the Termination Height value
                        if (bestSolution[0] == terminationHeight)
                        {
                            break;
                        }
                    }

                    // Check if the neighbor solution is already in tabu list
                    bool inTabu = false;

                    for (int t = 0; t < tabuCount; t++)
                    {
                        bool checkNextTabu = false;

                        for (int c = 1; c <= rectangleCount * 2; c++)
                        {
                            if (tabu[t][c] != neighbor[c])
                            {
                                // this neighbor is not this item
                                checkNextTabu = true;
                                break;
                            }
                        }

                        if (checkNextTabu)
                        {
                            continue;
                        }

                        // The order and rotation of tabu[t] is the same as neighbor's 
                        // This neighbor solution is already in tabu list
                        inTabu = true;
                    }

                    // If not in tabu list, record the neighbor solution
                    if (!inTabu)
                    {
                        if (bestNeighborIsEmpty)
                        {
                            // First admitted neighbor
                            Array.Copy(neighbor, bestNeighbor, rectangleCount * 2 + 1);
                            bestNeighborIsEmpty = false;
                        }
                        else
                        {
                            if (bestNeighbor[0] > neighbor[0])
                            {
                                Array.Copy(neighbor, bestNeighbor, rectangleCount * 2 + 1);
                            }
                        }
                    }
                }

                // There is no neighbor better than current solution
                if (currentSolution[0] <= bestNeighbor[0])
                {
                    Array.Copy(currentSolution, tabu[tabuCount], rectangleCount * 2 + 1);

                    if (tabuCount < tabuLength - 1)
                    {
                        tabuCount++;
                    }
                    else
                    {
                        tabuCount++;
                        break;
                    }
                }

                // Move to the best neighbor
                Array.Copy(bestNeighbor, currentSolution, rectangleCount * 2 + 1);
            }

            // Rerun bestSolution to generate the rectangles' positions saved in rectangle list
            bestSolution[0] = CalculateTotalHeight(rectangles, bestSolution, rectangleCount, regionAvailable, regionWidthLimit, regionHeightLimit);

            #endregion

            #region Draw Best Solution

            // Draw the best solution on the form
            for (int r = 1; r <= rectangleCount; r++)
            {
                int recNumber = bestSolution[r];

                System.Drawing.Rectangle shape = rectangles[recNumber];

                int x = shape.X;
                int y = shape.Y;
                int w = shape.Width;
                int h = shape.Height;

                System.Drawing.Rectangle shapeDraw;

                if (bestSolution[r + rectangleCount] == 1)
                {
                    // non-rotation
                    shapeDraw = new System.Drawing.Rectangle(x, y, w, h);
                }
                else
                {
                    // rotation
                    shapeDraw = new System.Drawing.Rectangle(x, y, h, w);
                }

                Graphics g2 = CreateGraphics();
                g2.FillRectangle(fill, shapeDraw);

                Graphics g1 = CreateGraphics();
                g1.DrawRectangle(border, shapeDraw);
            }

            #endregion

            #region Set Screen Info
            label1.Text = "M3d Calculation Done!";
            label2.Text = "Solution Compared:" + solutionCount.ToString();
            label3.Text = "Tabu:" + tabuCount.ToString();
            label4.Text = "Best Height:" + bestSolution[0].ToString();
            label7.Text = "Winforms' top left corner is (0,0). Will change picture orientation in the word report.";

            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            label8.Text = "Search Time: " + elapsedTime;
            #endregion
        }


        private bool[,] InitialRegion(int width, int height)
        {
            bool[,] regionAvailable = new bool[width + 1, height + 1];

            // Initial the internal region area to all true (available)
            // But the left and bottom frame to false (not available)
            for (int i = 0; i <= width; i++)
            {
                for (int j = 0; j <= height; j++)
                {
                    if (i == 0 || j == 0)
                    {
                        // Left or bottom region frame
                        regionAvailable[i, j] = false;
                    }
                    else
                    {
                        // Internal region area
                        regionAvailable[i, j] = true;
                    }
                }
            }

            return regionAvailable;
        }

        private int[] InitialRandomOrder(int count)
        {
            int[] arr = new int[count * 2 + 1];
            Random rnd = new Random();
            int tmp;

            for (int i = 0; i < arr.Length; i++)
            {
                if (i == 0)
                {
                    arr[i] = 0;
                }
                else if (i <= count)
                {
                    // Generate random order
                    tmp = rnd.Next(1,count + 1);
                    while (IsDup(tmp, arr))
                    {
                        tmp = rnd.Next(count + 1);
                    }
                    arr[i] = tmp;
                }
                else
                {
                    // 1 for non-rotation, 2 for rotation
                    // Initial soluton non-rotation
                    arr[i] = 1;
                }
            }

            return arr;
        }

        private bool IsDup(int tmp, int[] arr)
        {
            foreach(var item in arr)
            {
                if (item == tmp)
                {
                    return true;
                }
            }
            return false;
        }

        private int CalculateTotalHeight(System.Drawing.Rectangle[] rec, int[] HeightOrder, int count, bool[,] region, int widthLimit, int heightLimit)
        {
            // Declare total height
            int totalHeight = 0;

            // Movable flags for position calculation
            bool downMovable;
            bool leftMovable;

            // Loop for fill all rectangles by order
            for (int i = 1; i <= count; i++)
            {
                //totalHeight = 0;
                downMovable = true;
                leftMovable = true;

                int rectangleNumber = HeightOrder[i];
                int rotation = HeightOrder[i + count];
                int itemWidth;
                int itemHeight;
                if (rotation == 1)
                {
                    // non-rotation
                    itemWidth = rec[rectangleNumber].Width;
                    itemHeight = rec[rectangleNumber].Height;
                }
                else
                {
                    // rotation
                    itemWidth = rec[rectangleNumber].Height;
                    itemHeight = rec[rectangleNumber].Width;
                }
                

                // Calculate start position
                int positionX = widthLimit - itemWidth + 1;
                int positionY = heightLimit - itemHeight + 1;

                // Loop for calculate one rectangle's position
                while (downMovable || leftMovable)
                {
                    // Reset flags in While loop
                    downMovable = true;
                    leftMovable = true;

                    // Loop for down forward calculation
                    for (int h = positionY - 1; h >= 0; h--)
                    {
                        bool breakLoops = false;

                        for (int w = positionX + itemWidth - 1; w >= positionX; w--)
                        {
                            if (region[w, h] == false)
                            {
                                positionY = h + 1;
                                downMovable = false;

                                // End nested loops 
                                breakLoops = true;
                                break;
                            }
                        }

                        if (breakLoops)
                        {
                            break;
                        }
                    }

                    // Loop for left forward calculation
                    for (int w = positionX - 1; w >= 0; w--)
                    {
                        bool breakLoops = false;

                        int startPositionX;
                        startPositionX = positionX;

                        for (int h = positionY + itemHeight - 1; h >= positionY; h--)
                        {
                            if (region[w, h] == false)
                            {
                                positionX = w + 1;
                                leftMovable = false;
                                

                                // Moved left, need check downMovable again!
                                if (w != startPositionX - 1)
                                {
                                    downMovable = true;
                                }

                                // End nested loops
                                breakLoops = true;
                                break;
                            }
                        }

                        if (breakLoops)
                        {
                            break;
                        }
                    }

                    // Check if rectangle position is fixed
                    if (downMovable == false && leftMovable == false)
                    {
                        // "while loop" will be ended
                        // Record current position
                        rec[rectangleNumber].X = positionX;
                        rec[rectangleNumber].Y = positionY;

                        if (totalHeight < positionY + itemHeight)
                        {
                            totalHeight = positionY + itemHeight;
                        }

                        // Disable position area for other rectangles
                        for (int x = positionX; x < positionX + itemWidth; x++)
                        {
                            for (int y = positionY; y < positionY + itemHeight; y++)
                            {
                                region[x, y] = false;
                            }
                        }
                    }
                }
            }

            // Reset regionAvailabe
            //region = InitialRegion(widthLimit, heightLimit);
            // Initial the internal region area to all true (available)
            // But the left and bottom frame to false (not available)
            for (int i = 0; i <= widthLimit; i++)
            {
                for (int j = 0; j <= heightLimit; j++)
                {
                    if (i == 0 || j == 0)
                    {
                        // Left or bottom region frame
                        region[i, j] = false;
                    }
                    else
                    {
                        // Internal region area
                        region[i, j] = true;
                    }
                }
            }

            return totalHeight;
        }
    }
}
