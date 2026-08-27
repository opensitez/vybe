// vybe-test: csharp/csharp_functional_patterns/function_composition_applies_in_sequence
// origin: languages/csharp/tests/csharp/test_csharp_functional_patterns.rs

using static __Harness;

System.Func<int,int> triple=x=>x*3;
System.Func<int,int> addOne=x=>x+1;
var composed=new[]{1,2,3}
.Select(triple).Select(addOne);
foreach(var n in composed) __P((n).ToString());
__Check("4\n7\n10");

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
