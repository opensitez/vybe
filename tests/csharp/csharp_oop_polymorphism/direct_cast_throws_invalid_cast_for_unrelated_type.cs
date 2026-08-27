// vybe-test: csharp/csharp_oop_polymorphism/direct_cast_throws_invalid_cast_for_unrelated_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

using static __Harness;

string r="";
try{object o="hello"; int n=(int)o;}
catch(System.InvalidCastException){r="bad cast";}
__P((r).ToString());
__Check("bad cast");

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
