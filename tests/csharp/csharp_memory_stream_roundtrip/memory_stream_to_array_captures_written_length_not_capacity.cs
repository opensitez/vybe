// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_to_array_captures_written_length_not_capacity
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
stream.WriteByte(1);
stream.WriteByte(2);
var bytes = stream.ToArray();
__P((bytes.Length).ToString());
__P((bytes[1]).ToString());
__Check("2\n2");
