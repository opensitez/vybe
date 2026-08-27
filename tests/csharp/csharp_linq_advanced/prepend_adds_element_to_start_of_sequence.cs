// vybe-test: csharp/csharp_linq_advanced/prepend_adds_element_to_start_of_sequence
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

using static __Harness;

var result=new[]{2,3,4}
.Prepend(1);
__P((result.First()).ToString());
__Check("1");

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
