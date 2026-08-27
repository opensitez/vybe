// vybe-test: csharp/csharp_class_indexers/interface_indexer_is_invoked_through_interface_typed_reference
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

ICell row = new Row();
__P((row[1]).ToString());
__Check("b");

interface ICell {
    string this[int index] { get; }
}

class Row : ICell {
    string[] cells = { "a", "b" };
    public string this[int index] { get { return cells[index]; } }
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
