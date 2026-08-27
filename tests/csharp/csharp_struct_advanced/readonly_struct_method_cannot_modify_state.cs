// vybe-test: csharp/csharp_struct_advanced/readonly_struct_method_cannot_modify_state
// origin: languages/csharp/tests/csharp/test_csharp_struct_advanced.rs

using static __Harness;

var c=new Counter(5).Increment();
__P((c.Value).ToString());
__Check("6");

readonly struct Counter{
    public readonly int Value;
    public Counter(int v){Value=v;}
    public Counter Increment()=>new Counter(Value+1);
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
