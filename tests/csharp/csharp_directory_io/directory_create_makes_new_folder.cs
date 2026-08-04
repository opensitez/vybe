// vybe-test: csharp/csharp_directory_io/directory_create_makes_new_folder
// origin: languages/csharp/tests/csharp/test_csharp_directory_io.rs

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

string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vybe_test_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(path);
__P((System.IO.Directory.Exists(path)).ToString());
System.IO.Directory.Delete(path);
__Check("True");
