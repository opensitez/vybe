// vybe-test: csharp/csharp_io_compression_gzip_roundtrip/gzip_compression_case_13

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

var ms = new System.IO.MemoryStream();
using (var gz = new System.IO.Compression.GZipStream(ms, System.IO.Compression.CompressionMode.Compress, true)) {
    byte[] data = System.Text.Encoding.UTF8.GetBytes("GZipData_13");
    gz.Write(data, 0, data.Length);
}
__P((ms.Length > 0).ToString());
__Check("True");
