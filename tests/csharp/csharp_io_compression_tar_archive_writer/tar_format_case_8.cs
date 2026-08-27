// vybe-test: csharp/csharp_io_compression_tar_archive_writer/tar_format_case_8

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

var entryType = System.Formats.Tar.TarEntryType.RegularFile;
__P(((int)entryType).ToString());
__Check("48");
