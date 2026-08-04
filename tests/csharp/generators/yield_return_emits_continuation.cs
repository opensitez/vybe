// vybe-test: csharp/generators/yield_return_emits_continuation
// origin: languages/csharp/tests/csharp/test_generators.rs

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
    public static IEnumerable<int> Count() {
        yield return 1;
        yield return 2;
        yield return 3;
    }
    public static void Run() {
        var g = Count();
        __P((g).ToString());
    }
}
Program.Run();
__Check("[continuation]");
