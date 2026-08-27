// vybe-test: csharp/csharp_records_advanced/record_inheritance_preserves_base_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var dog = new Dog("Rex", 5);
__P((dog.Name).ToString());
__P((dog.Age).ToString());
__Check("Rex\n5");

record Animal(string Name);

record Dog(string Name, int Age) : Animal(Name);

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
