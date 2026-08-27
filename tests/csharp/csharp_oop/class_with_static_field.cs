// vybe-test: csharp/csharp_oop/class_with_static_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

var a = new Counter();
var b = new Counter();
var c = new Counter();
__P((Counter.Count).ToString());
__Check("3");

class Counter {
    public static int Count = 0;
    public Counter() { Count++; }
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
