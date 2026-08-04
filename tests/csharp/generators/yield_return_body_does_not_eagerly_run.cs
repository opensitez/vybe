// vybe-test: csharp/generators/yield_return_body_does_not_eagerly_run
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
    public static IEnumerable<int> Loud() {
        __P(("bad: body ran without resume").ToString());
        yield return 1;
    }
    public static void Run() {
        var _ = Loud();
        __P(("ok").ToString());
    }
}
Program.Run();
__Check("ok");
