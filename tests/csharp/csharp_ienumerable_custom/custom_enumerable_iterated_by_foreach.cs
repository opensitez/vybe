// vybe-test: csharp/csharp_ienumerable_custom/custom_enumerable_iterated_by_foreach
// origin: languages/csharp/tests/csharp/test_csharp_ienumerable_custom.rs

using static __Harness;

int sum=0;
foreach(var n in new UpTo(5)) sum+=n;
__P((sum).ToString());
__Check("15");

class UpTo:System.Collections.Generic.IEnumerable<int>{
    int _max;
    public UpTo(int max){_max=max;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=1;i<=_max;i++) yield return i;
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
