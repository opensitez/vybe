// vybe-test: csharp/csharp_list_filter_contracts/list_filter_contracts_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_list_filter_contracts.rs

using static __Harness;

// list_filter_contracts
string feature = "list_filter_contracts";
__P((feature.Substring(0, 1) == feature[0].ToString()).ToString());
__Check("True");

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
