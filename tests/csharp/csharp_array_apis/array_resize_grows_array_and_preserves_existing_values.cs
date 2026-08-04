// vybe-test: csharp/csharp_array_apis/array_resize_grows_array_and_preserves_existing_values
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

var values = new[] { 2, 4 }; System.Array.Resize(ref values, 4); foreach (var value in values) __P((value).ToString());
__Check("2\n4\n0\n0");
