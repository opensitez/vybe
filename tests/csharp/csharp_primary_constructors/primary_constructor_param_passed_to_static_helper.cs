// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_passed_to_static_helper
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Worker(3).Show()).ToString());
__Check("id=3");

class Worker(int id) {
    static string Format(int value) => "id=" + value;
    public string Show() => Format(id);
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
