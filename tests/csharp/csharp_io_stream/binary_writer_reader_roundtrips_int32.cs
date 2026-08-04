// vybe-test: csharp/csharp_io_stream/binary_writer_reader_roundtrips_int32
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
using(var bw = new System.IO.BinaryWriter(ms, System.Text.Encoding.UTF8, leaveOpen:true))
    bw.Write(12345);
ms.Position = 0;
using var br = new System.IO.BinaryReader(ms);
__P((br.ReadInt32()).ToString());
__Check("12345");
