// vybe-test: csharp/csharp_anonymous_types/anonymous_type_property_names_inferred_from_variable
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

using static __Harness;

int id=7;
string name="Bob";
var obj=new{id,name}
;
__P((obj.id).ToString());
__P((obj.name).ToString());
__Check("7\nBob");

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
