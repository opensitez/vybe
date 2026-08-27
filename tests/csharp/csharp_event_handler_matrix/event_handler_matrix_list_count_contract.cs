// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

using static __Harness;

// event_handler_matrix
var values = new System.Collections.Generic.List<int> { 77, 78, 77 }
;
__P((values.Count == 3).ToString());
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
