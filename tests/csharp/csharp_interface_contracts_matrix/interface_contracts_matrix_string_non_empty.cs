// vybe-test: csharp/csharp_interface_contracts_matrix/interface_contracts_matrix_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts_matrix.rs

using static __Harness;

// interface_contracts_matrix
string feature = "interface_contracts_matrix";
__P((feature.Length > 0).ToString());
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
