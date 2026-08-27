// vybe-test: csharp/csharp_oop_inheritance/is_operator_checks_runtime_type_in_hierarchy
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

using static __Harness;

object obj = new B();
__P((obj is A).ToString());
__P((obj is B).ToString());
__Check("True\nTrue");

class A { }

class B : A { }

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
