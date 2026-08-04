// vybe-test: csharp/csharp_file_io/file_copy_produces_identical_content
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

string src = System.IO.Path.GetTempFileName();
string dst = src + ".copy";
System.IO.File.WriteAllText(src, "data");
System.IO.File.Copy(src, dst, true);
__P((System.IO.File.ReadAllText(dst)).ToString());
System.IO.File.Delete(src);
System.IO.File.Delete(dst);
__Check("data");
