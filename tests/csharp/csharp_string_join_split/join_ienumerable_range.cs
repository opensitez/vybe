// vybe-test: csharp/csharp_string_join_split/join_ienumerable_range
// origin: languages/csharp/tests/csharp/test_csharp_string_join_split.rs

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

__P((string.Join("+",Enumerable.Range(1,4))).ToString());
__Check("1+2+3+4");
