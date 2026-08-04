// vybe-test: csharp/csharp_index_range_slice/range_both_indices_from_end
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

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

int[] data={10,20,30,40,50}; var slice=data[^4..^1]; __P((slice.Length).ToString()); __P((slice[0]).ToString()); __P((slice[2]).ToString());
__Check("3\n20\n40");
