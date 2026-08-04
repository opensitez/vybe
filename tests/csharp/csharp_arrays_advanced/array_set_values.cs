// vybe-test: csharp/csharp_arrays_advanced/array_set_values
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

var arr = new int[3];
arr[0] = 10;
arr[1] = 20;
arr[2] = 30;
__P((arr[0] + arr[1] + arr[2]).ToString());
__Check("60");
