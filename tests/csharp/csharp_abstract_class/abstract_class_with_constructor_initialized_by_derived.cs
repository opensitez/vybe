// vybe-test: csharp/csharp_abstract_class/abstract_class_with_constructor_initialized_by_derived
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class.rs

using static __Harness;

__P((new Tag("admin").Name).ToString());
__Check("admin");

abstract class Named{public string Name;public Named(string n){Name=n;}}

class Tag:Named{public Tag(string n):base(n){}}

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
