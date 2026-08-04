// vybe-test: csharp/csharp_io_stream/memory_stream_write_then_read_roundtrips_bytes
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
ms.WriteByte(42);
ms.Position = 0;
__P((ms.ReadByte()).ToString());
__Check("42");
