// vybe-test: csharp/csharp_constructor_patterns/primary_constructor_on_record_sets_all_fields
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

using static __Harness;

var p=new Person("Grace",40);
__P((p.Name).ToString());
__P((p.Age).ToString());
__Check("Grace\n40");

record Person(string Name,int Age);

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
