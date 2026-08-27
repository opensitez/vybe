// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_cast_to_int_array
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

int sum=0;
foreach(var v in System.Enum.GetValues(typeof(Score))) sum+=(int)v;
__P((sum).ToString());
__Check("9");

enum Score{A=1,B=3,C=5}

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
