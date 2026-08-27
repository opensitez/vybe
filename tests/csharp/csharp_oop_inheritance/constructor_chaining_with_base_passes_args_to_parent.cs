// vybe-test: csharp/csharp_oop_inheritance/constructor_chaining_with_base_passes_args_to_parent
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

__P((new Box("red").Color).ToString());
__Check("red");

class Shape { public string Color; public Shape(string c) { Color = c; } }

class Box : Shape { public Box(string c) : base(c) { } }

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
