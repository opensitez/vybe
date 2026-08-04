// vybe-test: csharp/linq_lambdas/func_as_return_value
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

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

Func<int, Func<int, int>> makeAdder = x => y => x + y;
var add5 = makeAdder(5);
__P((add5(3)).ToString());
__P((add5(10)).ToString());
__Check("8\n15");
