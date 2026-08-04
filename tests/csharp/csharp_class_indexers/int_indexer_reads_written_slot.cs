// vybe-test: csharp/csharp_class_indexers/int_indexer_reads_written_slot
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Buffer {
    int[] data = new int[3];
    public int this[int index] {
        get { return data[index]; }
        set { data[index] = value; }
    }
}
var buffer = new Buffer();
buffer[1] = 42;
__P((buffer[1]).ToString());
__Check("42");
