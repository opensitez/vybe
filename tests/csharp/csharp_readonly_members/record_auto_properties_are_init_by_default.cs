// vybe-test: csharp/csharp_readonly_members/record_auto_properties_are_init_by_default
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

using static __Harness;

var u=new User("Ada",20);
__P((u.Name).ToString());
__P((u.Age).ToString());
__Check("Ada\n20");

record User(string Name,int Age);

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
