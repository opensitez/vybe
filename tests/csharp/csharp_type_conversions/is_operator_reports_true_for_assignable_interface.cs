// vybe-test: csharp/csharp_type_conversions/is_operator_reports_true_for_assignable_interface
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

using static __Harness;
using System.Collections.Generic;

object item = new List<int>();
__P((item is IEnumerable<int>).ToString());
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
