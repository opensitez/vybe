// vybe-test: csharp/csharp_collection_initializer_syntax/object_initializer_can_invoke_properties
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;

var user = new User { Name = "Ada" }
;
__P((user.Name).ToString());
__Check("Ada");

class User { public string Name { get; set; } }

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
