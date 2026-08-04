// vybe-test: csharp/csharp_tuples_ranges/range_array_slice
// origin: languages/csharp/tests/csharp/test_csharp_tuples_ranges.rs

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

var arr = new[] { 0, 1, 2, 3, 4 };
var slice = arr[1..4];
foreach (var x in slice) __P((x).ToString());
__Check("1\n2\n3");
