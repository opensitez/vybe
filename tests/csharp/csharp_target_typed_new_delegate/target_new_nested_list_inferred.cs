// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_nested_list_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Collections.Generic.List<System.Collections.Generic.List<int>> grid = new();
System.Collections.Generic.List<int> row = new() { 1, 2 }
;
grid.Add(row);
__P((grid[0][1]).ToString());
__Check("2");

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
