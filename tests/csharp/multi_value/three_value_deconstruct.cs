// vybe-test: csharp/multi_value/three_value_deconstruct
// origin: languages/csharp/tests/csharp/test_multi_value.rs

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

class Program {
    public static (int, int, int) Rgb() {
        return (10, 20, 30);
    }
    public static void Run() {
        var (r, g, b) = Rgb();
        __P((r).ToString());
        __P((g).ToString());
        __P((b).ToString());
    }
}
Program.Run();
__Check("10\n20\n30");
