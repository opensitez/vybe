// vybe-test: csharp/common_patterns/params_with_normal_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

class Fmt {
    public static string Build(string prefix, params int[] nums) {
        return prefix + ": " + string.Join(",", nums);
    }
}
__P((Fmt.Build("nums", 1, 2, 3)).ToString());
__Check("nums: 1,2,3");
