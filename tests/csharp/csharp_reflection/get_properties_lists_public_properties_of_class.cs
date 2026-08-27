// vybe-test: csharp/csharp_reflection/get_properties_lists_public_properties_of_class
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

__P((typeof(Item).GetProperties().Length).ToString());
__Check("2");

class Item { public int Id {get;set;} public string Name {get;set;} }

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
