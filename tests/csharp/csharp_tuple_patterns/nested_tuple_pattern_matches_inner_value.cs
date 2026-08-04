// vybe-test: csharp/csharp_tuple_patterns/nested_tuple_pattern_matches_inner_value
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

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

var data=((1,2),(3,4));
var((a,b),(c,d))=data;
__P((a+b+c+d).ToString());
__Check("10");
