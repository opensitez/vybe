// vybe-test: csharp/csharp_design_patterns/strategy_pattern_swaps_algorithm_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

using static __Harness;

ISort s=new Ascending();
__P((string.Join(",",s.Sort(new[]{3,1,2}))).ToString());
s=new Descending();
__P((string.Join(",",s.Sort(new[]{3,1,2}))).ToString());
__Check("1,2,3\n3,2,1");

interface ISort{int[] Sort(int[] a);}

class Ascending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);return c;}}

class Descending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);System.Array.Reverse(c);return c;}}

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
