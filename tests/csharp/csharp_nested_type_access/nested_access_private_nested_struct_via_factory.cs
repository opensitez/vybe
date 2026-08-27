// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_struct_via_factory
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Builder().Build()).ToString());
__Check("8");

class Builder{struct Part{public int N;} Part Make(){return new Part{N=8};} public int Build()=>Make().N;}

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
