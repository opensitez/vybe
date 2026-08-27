// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_read_in_instance_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var c = new Counter(10);
c.Next();
c.Next();
__P((c.Value).ToString());
__Check("12");

class Counter(int start) {
    int current = start;
    public int Next() => ++current;
    public int Value => current;
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
