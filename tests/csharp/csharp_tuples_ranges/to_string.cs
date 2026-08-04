// vybe-test: csharp/csharp_tuples_ranges/to_string
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

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

int x = 42;
__P((x.ToString()).ToString());
__P((42.ToString()).ToString());
__Check("42\n42");
