// vybe-test: csharp/csharp_ranges_indices/index_from_end_one_is_last_element
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

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

int[] a={1,2,3,4,5}; __P((a[^1]).ToString());
__Check("5");
