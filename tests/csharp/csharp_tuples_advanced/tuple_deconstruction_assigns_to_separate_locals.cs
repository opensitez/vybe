// vybe-test: csharp/csharp_tuples_advanced/tuple_deconstruction_assigns_to_separate_locals
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

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

var (a, b, c) = (10, 20, 30);
__P((a+b+c).ToString());
__Check("60");
