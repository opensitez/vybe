// vybe-test: csharp/csharp_primary_constructors/primary_constructor_two_instances_are_independent
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var a = new Slot(1);
var b = new Slot(2);
__P((a.Id).ToString());
__P((b.Id).ToString());
__Check("1\n2");

class Slot(int id) { public int Id => id; }

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
