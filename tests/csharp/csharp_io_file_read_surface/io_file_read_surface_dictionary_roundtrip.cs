// vybe-test: csharp/csharp_io_file_read_surface/io_file_read_surface_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_io_file_read_surface.rs

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

// io_file_read_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[89] = 90; __P((map.ContainsKey(89) && map[89] == 90).ToString());
__Check("True");
