// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

using static __Harness;

// record_shape_basics
int seed = 39;
int[] numbers = new int[] { seed, seed + 1, seed + 2 }
;
__P((numbers.Length == 3).ToString());
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
