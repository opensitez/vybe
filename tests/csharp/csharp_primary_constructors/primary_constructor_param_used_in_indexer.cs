// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_indexer
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

var r = new Row(3);
r[1] = 9;
__P((r[1]).ToString());
__Check("9");

class Row(int size) {
    int[] cells = new int[size];
    public int this[int i] { get => cells[i]; set => cells[i] = value; }
}

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
