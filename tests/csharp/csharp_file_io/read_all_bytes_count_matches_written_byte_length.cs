// vybe-test: csharp/csharp_file_io/read_all_bytes_count_matches_written_byte_length
// origin: languages/csharp/tests/csharp/test_csharp_file_io.rs

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

string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllBytes(path, new byte[]{1,2,3,4,5});
__P((System.IO.File.ReadAllBytes(path).Length).ToString());
System.IO.File.Delete(path);
__Check("5");
