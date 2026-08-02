// vybe-test: csharp/csharp_design_patterns/strategy_pattern_swaps_algorithm_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ISort{int[] Sort(int[] a);}
class Ascending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);return c;}}
class Descending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);System.Array.Reverse(c);return c;}}
ISort s=new Ascending();
__Check((string.Join(",",s.Sort(new[]{3,1,2}))).ToString(), "1,2,3");
s=new Descending();
__Check((string.Join(",",s.Sort(new[]{3,1,2}))).ToString(), "3,2,1");
