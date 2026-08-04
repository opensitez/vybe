// vybe-test: csharp/csharp_design_patterns/strategy_pattern_swaps_algorithm_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface ISort{int[] Sort(int[] a);}
class Ascending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);return c;}}
class Descending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);System.Array.Reverse(c);return c;}}
ISort s=new Ascending();
__P((string.Join(",",s.Sort(new[]{3,1,2}))).ToString());
s=new Descending();
__P((string.Join(",",s.Sort(new[]{3,1,2}))).ToString());
__Check("1,2,3\n3,2,1");
