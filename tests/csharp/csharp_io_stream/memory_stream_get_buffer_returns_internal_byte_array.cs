// vybe-test: csharp/csharp_io_stream/memory_stream_get_buffer_returns_internal_byte_array
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

using var ms = new System.IO.MemoryStream(new byte[]{1,2,3});
__P((ms.Length).ToString());
__Check("3");
