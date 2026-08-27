// vybe-test: csharp/csharp_enum_metaprogramming/enum_switch_on_parsed_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

var m=(Mode)System.Enum.Parse(typeof(Mode),"On");
string s=m==Mode.On?"yes":"no";
__P((s).ToString());
__Check("yes");

enum Mode{On,Off}

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
