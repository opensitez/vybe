// vybe-test: csharp/csharp_deconstruction/tuple_deconstruction_assigns_two_scalars
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

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

var (x, y) = (3, 4);
__P((x).ToString());
__P((y).ToString());
__Check("3\n4");
