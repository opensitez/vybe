// vybe-test: csharp/csharp_goto_switch_labels/goto_case_on_enum_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

Color c = Color.Red;
string name = "";
switch (c) {
    case Color.Red: name += "R"; goto case Color.Green;
    case Color.Green: name += "G"; break;
    case Color.Blue: name += "B"; break;
}
__P((name).ToString());
__Check("RG");

enum Color { Red, Green, Blue }

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
