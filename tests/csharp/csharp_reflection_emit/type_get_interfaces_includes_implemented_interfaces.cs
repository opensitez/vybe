// vybe-test: csharp/csharp_reflection_emit/type_get_interfaces_includes_implemented_interfaces
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

using static __Harness;

bool has=System.Array.Exists(typeof(Foo).GetInterfaces(),t=>t==typeof(IFoo));
__P((has).ToString());
__Check("True");

interface IFoo{}

class Foo:IFoo{}

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
