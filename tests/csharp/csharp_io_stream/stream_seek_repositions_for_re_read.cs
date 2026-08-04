// vybe-test: csharp/csharp_io_stream/stream_seek_repositions_for_re_read
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
ms.WriteByte(7);
ms.Seek(0, System.IO.SeekOrigin.Begin);
__P((ms.ReadByte()).ToString());
__Check("7");
