// vybe-test: csharp/csharp_tuples_advanced/tuple_with_eight_elements_uses_rest_field
// origin: languages/csharp/tests/csharp/test_csharp_tuples_advanced.rs

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

var t = (1,2,3,4,5,6,7,8);
__P((t.Item8).ToString());
__Check("8");
