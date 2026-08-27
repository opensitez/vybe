// vybe-test: csharp/csharp_linq_chaining/zip_three_sequences_pairwise_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

using static __Harness;

var a=new[]{1,2,3}
;
var b=new[]{10,20,30}
;
var result=a.Zip(b).Select(t=>t.First+t.Second);
__P((string.Join(",",result)).ToString());
__Check("11,22,33");

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
