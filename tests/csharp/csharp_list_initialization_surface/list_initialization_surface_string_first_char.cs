// vybe-test: csharp/csharp_list_initialization_surface/list_initialization_surface_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_list_initialization_surface.rs

using static __Harness;

// list_initialization_surface
string feature = "list_initialization_surface";
__P((feature[0] == feature[0]).ToString());
__Check("True");

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
