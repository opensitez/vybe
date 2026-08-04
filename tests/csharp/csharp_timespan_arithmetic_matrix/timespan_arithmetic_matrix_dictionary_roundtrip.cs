// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

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

// timespan_arithmetic_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[95] = 96; __P((map.ContainsKey(95) && map[95] == 96).ToString());
__Check("True");
