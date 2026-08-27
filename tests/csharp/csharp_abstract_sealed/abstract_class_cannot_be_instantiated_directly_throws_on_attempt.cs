// vybe-test: csharp/csharp_abstract_sealed/abstract_class_cannot_be_instantiated_directly_throws_on_attempt
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

using static __Harness;

string result = "ok";
try {
    var obj = System.Activator.CreateInstance(typeof(Base));
    result = "created";
}
catch (System.MemberAccessException) {
    result = "blocked";
}
catch (System.Exception) {
    result = "blocked";
}
__P((result).ToString());
__Check("blocked");

abstract class Base { }

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
