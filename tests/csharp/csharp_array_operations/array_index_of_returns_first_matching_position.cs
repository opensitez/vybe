// vybe-test: csharp/csharp_array_operations/array_index_of_returns_first_matching_position
// origin: languages/csharp/tests/csharp/test_csharp_array_operations.rs

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

string[] a = {"a","b","c","b"};
__P((System.Array.IndexOf(a,"b")).ToString());
__Check("1");
