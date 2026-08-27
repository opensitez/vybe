// vybe-test: csharp/csharp_access_modifiers/private_setter_means_field_read_only_from_outside
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

using static __Harness;

var c=new Counter();
c.Tick();
c.Tick();
__P((c.Count).ToString());
__Check("2");

class Counter{
    public int Count{get;private set;}
    public void Tick(){Count++;}
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
