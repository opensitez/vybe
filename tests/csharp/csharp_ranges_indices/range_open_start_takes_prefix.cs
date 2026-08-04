// vybe-test: csharp/csharp_ranges_indices/range_open_start_takes_prefix
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

int[] a={1,2,3,4,5}; var s=a[..3];
__P((s.Length).ToString()); __P((s[2]).ToString());
__Check("3\n3");
