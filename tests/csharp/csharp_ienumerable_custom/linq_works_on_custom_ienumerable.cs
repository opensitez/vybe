// vybe-test: csharp/csharp_ienumerable_custom/linq_works_on_custom_ienumerable
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

using static __Harness;

__P((new Odds(4).Sum()).ToString());
__Check("16");

class Odds:System.Collections.Generic.IEnumerable<int>{
    int _count;
    public Odds(int count){_count=count;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=0;i<_count;i++) yield return 2*i+1;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}

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
