// vybe-test: csharp/csharp_constructor_chains/record_primary_constructor_members_are_available_immediately
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

var user = new User("Ada", 20);
__P((user.Name).ToString());
__P((user.Age).ToString());
__Check("Ada\n20");

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
