// vybe-test: csharp/csharp_deconstruction/deconstruction_with_existing_variables_reassigns_them
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction.rs

using static __Harness;

int first = 0;
int second = 0;
(first, second) = (7, 9);
__P((first).ToString());
__P((second).ToString());
__Check("7\n9");

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
