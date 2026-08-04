// vybe-test: csharp/csharp_index_from_end/range_to_end_from_index_from_end_produces_tail_slice
// origin: languages/csharp/tests/csharp/test_csharp_index_from_end.rs

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

int[] data = { 1, 2, 3, 4 };
var tail = data[2..^0];
__P((tail.Length).ToString());
__P((tail[0]).ToString());
__Check("2\n3");
