// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_static_factory_method
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((Pool.Token.Make(21).Id).ToString());
__Check("21");

class Pool{public class Token{public int Id; public static Token Make(int id)=>new Token{Id=id};}}

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
