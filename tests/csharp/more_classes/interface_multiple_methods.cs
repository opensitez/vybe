// vybe-test: csharp/more_classes/interface_multiple_methods
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICalc {
            int Add(int a, int b);
            int Mul(int a, int b);
        }
        class Calc : ICalc {
            public int Add(int a, int b) { return a + b; }
            public int Mul(int a, int b) { return a * b; }
        }
        var c = new Calc();
        __Check((c.Add(3, 4)).ToString(), "7");
        __Check((c.Mul(3, 4)).ToString(), "12");
