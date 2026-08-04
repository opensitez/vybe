// vybe-test: csharp/linq_lambdas/linq_any_all
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

var nums = new List<int> { 2, 4, 6, 8 };
__P((nums.All(x => x % 2 == 0)).ToString());
__P((nums.Any(x => x > 5)).ToString());
__P((nums.Any(x => x > 10)).ToString());
__Check("True\nTrue\nFalse");
