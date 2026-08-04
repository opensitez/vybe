// vybe-test: csharp/csharp_array_apis/array_clear_resets_values_to_default
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

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

var values = new[] { 1, 2, 3 }; System.Array.Clear(values, 1, 2); foreach (var value in values) __P((value).ToString());
__Check("1\n0\n0");
