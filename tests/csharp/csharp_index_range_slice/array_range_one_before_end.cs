// vybe-test: csharp/csharp_index_range_slice/array_range_one_before_end
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

int[] data={2,4,6,8,10}; var slice=data[..^1]; __P((slice[slice.Length-1]).ToString());
__Check("8");
