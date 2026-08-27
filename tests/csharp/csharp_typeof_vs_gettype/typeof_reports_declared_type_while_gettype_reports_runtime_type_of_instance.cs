// vybe-test: csharp/csharp_typeof_vs_gettype/typeof_reports_declared_type_while_gettype_reports_runtime_type_of_instance
// origin: languages/csharp/tests/csharp/test_csharp_typeof_vs_gettype.rs

using static __Harness;

Animal pet = new Dog();
__P((typeof(Animal).Name).ToString());
__P((pet.GetType().Name).ToString());
__Check("Animal\nDog");

class Animal { }

class Dog : Animal { }

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
