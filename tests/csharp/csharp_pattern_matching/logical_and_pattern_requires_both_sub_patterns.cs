// vybe-test: csharp/csharp_pattern_matching/logical_and_pattern_requires_both_sub_patterns
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

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

int n = 15;
__P((n is > 10 and < 20).ToString());
__Check("True");
