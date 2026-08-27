// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_indexer_reads_values_through_interface
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

IReadIndex words = new Words();
__P((words[0]).ToString());
__P((words[1]).ToString());
__Check("alpha\nbeta");

interface IReadIndex { string this[int index] { get; } }

class Words : IReadIndex {
    string[] values = new[] { "alpha", "beta" };
    string IReadIndex.this[int index] { get { return values[index]; } }
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
