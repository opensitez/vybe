// vybe-test: csharp/csharp_null_propagation/null_coalescing_can_select_new_object_instance
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;

Box box = null;
box ??= new Box { Name = "created" }
;
__P((box.Name).ToString());
__Check("created");

class Box { public string Name; }

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
