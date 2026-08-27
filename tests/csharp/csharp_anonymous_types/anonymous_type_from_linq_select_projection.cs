// vybe-test: csharp/csharp_anonymous_types/anonymous_type_from_linq_select_projection
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

using static __Harness;

var data=new[]{(Id:1,Name:"a"),(Id:2,Name:"b")}
;
var result=data.Select(d=>new{d.Id,Upper=d.Name.ToUpper()}).ToList();
__P((result[1].Upper).ToString());
__Check("B");

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
