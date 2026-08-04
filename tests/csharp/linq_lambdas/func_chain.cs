// vybe-test: csharp/linq_lambdas/func_chain
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

Func<int, int> doubleIt = x => x * 2;
Func<int, int> addOne = x => x + 1;
__P((addOne(doubleIt(5))).ToString());
__P((doubleIt(addOne(5))).ToString());
__Check("11\n12");
