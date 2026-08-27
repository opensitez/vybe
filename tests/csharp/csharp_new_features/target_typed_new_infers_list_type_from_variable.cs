// vybe-test: csharp/csharp_new_features/target_typed_new_infers_list_type_from_variable
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

using static __Harness;

System.Collections.Generic.List<int> nums = new();
nums.Add(1);
nums.Add(2);
__P((nums.Count).ToString());
__Check("2");

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
