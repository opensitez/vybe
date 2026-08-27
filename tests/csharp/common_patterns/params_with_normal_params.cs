// vybe-test: csharp/common_patterns/params_with_normal_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P((Fmt.Build("nums", 1, 2, 3)).ToString());
__Check("nums: 1,2,3");

class Fmt {
    public static string Build(string prefix, params int[] nums) {
        return prefix + ": " + string.Join(",", nums);
    }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
