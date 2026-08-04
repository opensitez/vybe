// vybe-test: csharp/csharp_delegates/lambda_block_body
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

Func<int, string> classify = x => {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
};
__P((classify(5)).ToString());
__P((classify(-3)).ToString());
__P((classify(0)).ToString());
__Check("positive\nnegative\nzero");
