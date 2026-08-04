// vybe-test: csharp/csharp_linq_aggregate_element/min_by_max_by_same_sequence_lengths
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

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

var words=new[]{"go","stop","run"};
__P((words.MinBy(w=>w.Length).Length).ToString());
__P((words.MaxBy(w=>w.Length).Length).ToString());
__Check("2\n4");
