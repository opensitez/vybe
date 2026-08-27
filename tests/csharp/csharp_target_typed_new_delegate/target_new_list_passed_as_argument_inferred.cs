// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_passed_as_argument_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

int Sum(System.Collections.Generic.List<int> xs) { int s = 0; foreach (var x in xs) s += x; return s; }
System.Collections.Generic.List<int> data = new() { 1, 2, 3 }
;
__P((Sum(data)).ToString());
__Check("6");

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
