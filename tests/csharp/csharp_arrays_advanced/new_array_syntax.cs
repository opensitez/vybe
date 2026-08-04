// vybe-test: csharp/csharp_arrays_advanced/new_array_syntax
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

var arr = new[] { 10, 20, 30, 40, 50 };
__P((arr.Length).ToString());
__P((arr[2]).ToString());
__Check("5\n30");
