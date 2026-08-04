// vybe-test: csharp/csharp_scope_variables/var_keyword_infers_type_from_right_hand_side
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

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

var text = "hello";
var number = 42;
__P((text.GetType().Name).ToString());
__P((number.GetType().Name).ToString());
__Check("String\nInt32");
