// vybe-test: csharp/csharp_arrays_advanced/array_for_loop
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

var arr = new[] { 1, 2, 3, 4, 5 };
int sum = 0;
for (int i = 0; i < arr.Length; i++) {
    sum += arr[i];
}
__P((sum).ToString());
__Check("15");
