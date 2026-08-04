// vybe-test: csharp/advanced/array_foreach_sum
// origin: languages/csharp/tests/csharp/test_advanced.rs

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

var nums = new int[] { 10, 20, 30, 40 };
        var total = 0;
        foreach (var n in nums) { total = total + n; }
        __P((total).ToString());
__Check("100");
