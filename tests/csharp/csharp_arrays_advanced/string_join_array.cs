// vybe-test: csharp/csharp_arrays_advanced/string_join_array
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

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

var arr = new[] { "a", "b", "c" };
__P((string.Join(",", arr)).ToString());
__P((string.Join(" - ", arr)).ToString());
__Check("a,b,c\na - b - c");
