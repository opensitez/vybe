// vybe-test: csharp/more_classes/interface_multiple_methods
// origin: languages/csharp/tests/csharp/test_more_classes.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
        __P((c.Add(3, 4)).ToString());
        __P((c.Mul(3, 4)).ToString());
__Check("7\n12");
