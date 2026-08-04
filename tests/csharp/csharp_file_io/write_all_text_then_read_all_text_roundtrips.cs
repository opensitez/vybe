// vybe-test: csharp/csharp_file_io/write_all_text_then_read_all_text_roundtrips
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
System.IO.File.WriteAllText(path, "hello");
__P((System.IO.File.ReadAllText(path)).ToString());
System.IO.File.Delete(path);
__Check("hello");
