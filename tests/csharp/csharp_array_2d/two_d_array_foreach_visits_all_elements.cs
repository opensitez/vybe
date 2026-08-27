// vybe-test: csharp/csharp_array_2d/two_d_array_foreach_visits_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

using static __Harness;

int[,] m={{1,2},{3,4}}
;
int sum=0;
foreach(int n in m) sum+=n;
__P((sum).ToString());
__Check("10");

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
