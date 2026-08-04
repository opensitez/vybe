// vybe-test: csharp/csharp_convert_uri_path/path_change_extension_replaces_trailing_suffix
// origin: languages/csharp/tests/csharp/test_csharp_convert_uri_path.rs

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

// convert_uri_path
__P((System.IO.Path.ChangeExtension("data.txt", ".json")).ToString());
__Check("data.json");
