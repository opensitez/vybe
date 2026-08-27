// vybe-test: csharp/csharp_local_function_static/local_function_switch_expression
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

string Label(int n){string L(int x)=>x switch{1=>"one",2=>"two",_=>"other"}; return L(n);}
__P((Label(2)).ToString());
__Check("two");

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
