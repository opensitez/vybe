// vybe-test: csharp/common_patterns/params_with_normal_params
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Fmt {
    public static string Build(string prefix, params int[] nums) {
        return prefix + ": " + string.Join(",", nums);
    }
}
__Check((Fmt.Build("nums", 1, 2, 3)).ToString(), "nums: 1,2,3");
