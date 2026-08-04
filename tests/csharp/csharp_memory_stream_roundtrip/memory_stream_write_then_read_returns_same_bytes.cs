// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_write_then_read_returns_same_bytes
// origin: languages/csharp/tests/csharp/test_csharp_memory_stream_roundtrip.rs

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

using System.IO;
var stream = new MemoryStream();
var writer = new StreamWriter(stream);
writer.Write("payload");
writer.Flush();
stream.Position = 0;
var reader = new StreamReader(stream);
__P((reader.ReadToEnd()).ToString());
__Check("payload");
