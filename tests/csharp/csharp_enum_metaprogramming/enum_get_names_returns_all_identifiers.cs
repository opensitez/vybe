// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_returns_all_identifiers
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

using static __Harness;

foreach(var name in System.Enum.GetNames(typeof(Coin))) __P((name).ToString());
__Check("Penny\nNickel\nDime");

enum Coin{Penny,Nickel,Dime}

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
