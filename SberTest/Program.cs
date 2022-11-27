using System;

namespace SberTest
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            switch (args[0])
            {
                case "1":
                {
                    var solver = new FirstProblemSolver("кофе", 5);
                    solver.Solve();
                    break;
                }
                case "2":
                {
                    var solver = new SecondProblemSolver("output.txt", "этот текст нужно вставить");
                    solver.Solve();
                    break;
                }
                case "3":
                {
                    var solver = new ThirdProblemSolver();
                    solver.Solve();
                    break;
                }
                default:
                    break;
            }
        }
    }
}
