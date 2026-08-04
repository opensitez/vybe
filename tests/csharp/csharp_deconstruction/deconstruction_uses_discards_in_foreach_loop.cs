// vybe-test: csharp/csharp_deconstruction/deconstruction_uses_discards_in_foreach_loop
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

var items = new[] { ("x", 1), ("y", 2), ("z", 3) };
foreach (var (_, number) in items) {
    __P((number * 10).ToString());
}
__Check("10\n20\n30");
