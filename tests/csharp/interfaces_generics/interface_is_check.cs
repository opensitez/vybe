// vybe-test: csharp/interfaces_generics/interface_is_check
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

object b = new Bird();
object f = new Fish();
__P((b is IFlyable).ToString());
__P((f is IFlyable).ToString());
__Check("True\nFalse");

interface IFlyable { }

class Bird : IFlyable { }

class Fish { }

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
