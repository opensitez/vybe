// vybe-test: csharp/csharp_io_stream/stream_writer_reader_roundtrip_text_line
// origin: languages/csharp/tests/csharp/test_csharp_io_stream.rs

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

using var ms = new System.IO.MemoryStream();
using(var sw = new System.IO.StreamWriter(ms, leaveOpen:true)) sw.WriteLine("hello");
ms.Position = 0;
using var sr = new System.IO.StreamReader(ms);
__P((sr.ReadLine()).ToString());
__Check("hello");
