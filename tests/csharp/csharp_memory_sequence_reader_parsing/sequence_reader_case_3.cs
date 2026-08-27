// vybe-test: csharp/csharp_memory_sequence_reader_parsing/sequence_reader_case_3

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

var seq = new System.Buffers.ReadOnlySequence<byte>(new byte[] { (byte)3, 20, 30 });
var reader = new System.Buffers.SequenceReader<byte>(seq);
bool ok = reader.TryRead(out byte b);
__P(ok.ToString());
__P(b.ToString());
__Check("True\n3");
