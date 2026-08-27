// vybe-test: csharp/csharp_records_advanced/record_clone_via_with_does_not_mutate_original_instance
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var before = new User("Ada", 30);
var after = before with { Name = "Grace" }
;
__P((before.Name).ToString());
__P((after.Name).ToString());
__Check("Ada\nGrace");

record User(string Name, int Age);

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
