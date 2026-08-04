// vybe-test: csharp/csharp_pattern_matching/tuple_pattern_deconstructs_two_element_tuple
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

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

var point = (1, 0);
string axis = point switch {
    (0, 0) => "origin",
    (_, 0) => "x-axis",
    (0, _) => "y-axis",
    _       => "other"
};
__P((axis).ToString());
__Check("x-axis");
