// vybe-test: csharp/csharp_delegates/func_delegate
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

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

Func<int, int> square = x => x * x;
__P((square(5)).ToString());
__P((square(8)).ToString());
__Check("25\n64");
