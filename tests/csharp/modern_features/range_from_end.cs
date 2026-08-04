// vybe-test: csharp/modern_features/range_from_end
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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

int[] nums = { 1, 2, 3, 4, 5 };
int[] last3 = nums[^3..];
foreach (var n in last3) __P((n).ToString());
__Check("3\n4\n5");
