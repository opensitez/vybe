// vybe-test: csharp/csharp_memory_stream_roundtrip/memory_stream_seek_begin_repositions_read_cursor_for_second_pass
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
writer.Write("ab");
writer.Flush();
stream.Seek(0, SeekOrigin.Begin);
__P((stream.ReadByte()).ToString());
stream.Seek(0, SeekOrigin.Begin);
__P((stream.ReadByte()).ToString());
__Check("97\n97");
