// vybe-test: csharp/csharp_memory_array_buffer_writer_spans/buffer_writer_case_9

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var writer = new System.Buffers.ArrayBufferWriter<byte>();
var span = writer.GetSpan(4);
span[0] = (byte)9;
writer.Advance(1);
__P(writer.WrittenCount.ToString());
__P(writer.WrittenSpan[0].ToString());
__Check("1\n9");
