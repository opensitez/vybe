// vybe-test: csharp/csharp_null_handling/null_conditional_chains_through_nested_properties
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

using static __Harness;

Node head = null;
__P((head?.Next?.Value ?? -1).ToString());
__Check("-1");

class Node { public Node Next; public int Value; }

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
