// vybe-test: csharp/csharp_structs_value_semantics/struct_can_implement_generic_interface
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

IBox<int> box = new NumberBox();
__P((box.Read()).ToString());
__Check("14");

interface IBox<T> { T Read(); }

struct NumberBox : IBox<int> { public int Read() { return 14; } }

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
