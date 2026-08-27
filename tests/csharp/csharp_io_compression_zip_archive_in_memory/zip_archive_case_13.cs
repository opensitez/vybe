// vybe-test: csharp/csharp_io_compression_zip_archive_in_memory/zip_archive_case_13

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
using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true)) {
    var entry = zip.CreateEntry("file_13.txt");
    using var writer = new System.IO.StreamWriter(entry.Open());
    writer.Write("Content_13");
}
__P((ms.Length > 0).ToString());
__Check("True");
