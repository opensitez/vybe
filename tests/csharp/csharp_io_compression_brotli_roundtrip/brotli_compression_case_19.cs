// vybe-test: csharp/csharp_io_compression_brotli_roundtrip/brotli_compression_case_19

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

byte[] data = System.Text.Encoding.UTF8.GetBytes("BrotliData_19");
byte[] compressed = new byte[100];
bool ok = System.IO.Compression.BrotliEncoder.TryCompress(data, compressed, out int bytesWritten);
__P(ok.ToString());
__P((bytesWritten > 0).ToString());
__Check("True\nTrue");
