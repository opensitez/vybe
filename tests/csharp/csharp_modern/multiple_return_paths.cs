// vybe-test: csharp/csharp_modern/multiple_return_paths
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

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

string Classify(int x) {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
}
__P((Classify(5)).ToString());
__P((Classify(-3)).ToString());
__P((Classify(0)).ToString());
__Check("positive\nnegative\nzero");
