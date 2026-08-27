// vybe-test: csharp/csharp_local_function_static/local_function_string_builder_capture
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

string Join(int a,int b){var sb=new System.Text.StringBuilder(); string Append(int x){sb.Append(x); return sb.ToString();} Append(a); return Append(b);}
__P((Join(1,2)).ToString());
__Check("12");

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
