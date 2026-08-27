// vybe-test: csharp/csharp_record_struct/record_struct_copy_is_independent_value_after_mutation
// origin: languages/csharp/tests/csharp/test_csharp_record_struct.rs

using static __Harness;

var a=new Count(5);
var b=a;
b=b with{N=99}
;
__P((a.N).ToString());
__Check("5");

record struct Count(int N);

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
