// vybe-test: csharp/csharp_deconstruction/foreach_deconstruction_reads_tuple_sequence
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

var pairs = new[] { ("a", 1), ("b", 2) };
foreach (var (letter, number) in pairs) {
    __P((letter + number).ToString());
}
__Check("a1\nb2");
