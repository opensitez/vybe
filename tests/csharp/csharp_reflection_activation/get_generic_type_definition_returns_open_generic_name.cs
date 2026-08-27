// vybe-test: csharp/csharp_reflection_activation/get_generic_type_definition_returns_open_generic_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;

__P((typeof(Box<int>).GetGenericTypeDefinition().Name.Contains("Box")).ToString());
__Check("True");

class Box<T> { }

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
