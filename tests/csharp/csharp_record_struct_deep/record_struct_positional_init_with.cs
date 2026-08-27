// vybe-test: csharp/csharp_record_struct_deep/record_struct_positional_init_with
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

using static __Harness;

var u=new User("Ada"){Age=30}
;
var v=u with{Age=31}
;
__P((u.Age).ToString());
__P((v.Age).ToString());
__Check("30\n31");

record struct User(string Name){public int Age{get;init;}}

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
