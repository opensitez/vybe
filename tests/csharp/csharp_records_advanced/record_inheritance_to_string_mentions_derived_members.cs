// vybe-test: csharp/csharp_records_advanced/record_inheritance_to_string_mentions_derived_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

__P((new Cat("Milo", "Black").ToString().Contains("Color = Black")).ToString());
__Check("True");

record Animal(string Name);

record Cat(string Name, string Color) : Animal(Name);

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
