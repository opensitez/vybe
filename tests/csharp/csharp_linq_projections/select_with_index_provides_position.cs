// vybe-test: csharp/csharp_linq_projections/select_with_index_provides_position
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

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

var result = new[]{"a","b","c"}.Select((x,i) => $"{i}:{x}");
foreach(var s in result) __P((s).ToString());
__Check("0:a\n1:b\n2:c");
