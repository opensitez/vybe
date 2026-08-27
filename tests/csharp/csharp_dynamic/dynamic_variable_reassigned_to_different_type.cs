// vybe-test: csharp/csharp_dynamic/dynamic_variable_reassigned_to_different_type
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

using static __Harness;

dynamic v=42;
__P((v.GetType().Name).ToString());
v="hello";
__P((v.GetType().Name).ToString());
__Check("Int32\nString");

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
