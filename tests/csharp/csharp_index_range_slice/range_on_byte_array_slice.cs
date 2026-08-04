// vybe-test: csharp/csharp_index_range_slice/range_on_byte_array_slice
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

byte[] data={10,20,30,40}; var slice=data[1..3]; __P((slice[0]).ToString()); __P((slice[1]).ToString());
__Check("20\n30");
