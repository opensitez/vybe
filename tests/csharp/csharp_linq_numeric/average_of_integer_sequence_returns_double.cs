// vybe-test: csharp/csharp_linq_numeric/average_of_integer_sequence_returns_double
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

using static __Harness;

double avg=new[]{1,2,3,4,5}
.Average();
__P((avg).ToString());
__Check("3");

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
