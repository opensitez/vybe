// vybe-test: csharp/csharp_goto_switch_labels/goto_label_from_inner_if_in_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

using static __Harness;

int sum = 0;
for (int i = 0; i < 5; i++) {
    if (i == 3) goto done;
    sum += i;
}
done:
__P((sum).ToString());
__Check("3");

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
