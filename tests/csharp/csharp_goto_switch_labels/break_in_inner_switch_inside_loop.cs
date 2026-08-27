// vybe-test: csharp/csharp_goto_switch_labels/break_in_inner_switch_inside_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            switch (i) {
                case 0: log += "in"; break;
            }
            log += ";";
            break;
        case 1: log += "out"; break;
    }
}
__P((log).ToString());
__Check("in;out");

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
