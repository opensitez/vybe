// vybe-test: csharp/csharp_class_indexers/int_indexer_reads_written_slot
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((buffer[1]).ToString(), "42");
