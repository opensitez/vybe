// vybe-test: csharp/collections_advanced/array_indexof
// origin: languages/csharp/tests/csharp/test_collections_advanced.rs

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

string[] arr = { "a", "b", "c", "d" };
__P((Array.IndexOf(arr, "c")).ToString());
__P((Array.IndexOf(arr, "z")).ToString());
__Check("2\n-1");
