// vybe-test: csharp/csharp_pattern_matching_advanced/object_type_pattern_matches_base_class_instance
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

object pet = new Dog();
__P((pet is Animal).ToString());
__Check("True");

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
