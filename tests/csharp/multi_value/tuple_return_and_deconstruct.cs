// vybe-test: csharp/multi_value/tuple_return_and_deconstruct
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
    public static (int, int) Swap(int a, int b) {
        return (b, a);
    }
    public static void Run() {
        var (x, y) = Swap(1, 2);
        __P((x).ToString());
        __P((y).ToString());
    }
}
Program.Run();
__Check("2\n1");
