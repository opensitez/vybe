// vybe-test: csharp/csharp_class_indexers/int_indexer_reads_written_slot
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

using static __Harness;

var buffer = new Buffer();
buffer[1] = 42;
__P((buffer[1]).ToString());
__Check("42");

class Buffer {
    int[] data = new int[3];
    public int this[int index] {
        get { return data[index]; }
        set { data[index] = value; }
    }
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
