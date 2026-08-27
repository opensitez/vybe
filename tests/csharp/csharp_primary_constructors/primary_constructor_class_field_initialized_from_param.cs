// vybe-test: csharp/csharp_primary_constructors/primary_constructor_class_field_initialized_from_param
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Holder(99).Read()).ToString());
__Check("99");

class Holder(int seed) { int value = seed; public int Read() => value; }

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
