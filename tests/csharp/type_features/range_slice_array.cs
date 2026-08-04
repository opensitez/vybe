// vybe-test: csharp/type_features/range_slice_array
// origin: languages/csharp/tests/csharp/test_type_features.rs

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

var arr = new int[] { 10, 20, 30, 40, 50 };
        var sub = arr[1..3];
        __P((sub[0]).ToString());
        __P((sub[1]).ToString());
__Check("20\n30");
