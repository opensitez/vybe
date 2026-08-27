// vybe-test: csharp/csharp_records_advanced/with_expression_copies_record_and_changes_one_member
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var user = new User("Ada", 20);
var updated = user with { Age = 21 }
;
__P((user.Age).ToString());
__P((updated.Age).ToString());
__Check("20\n21");

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
