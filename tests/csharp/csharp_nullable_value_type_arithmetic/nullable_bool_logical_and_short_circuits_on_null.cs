// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_bool_logical_and_short_circuits_on_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_type_arithmetic.rs

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

bool? t = true;
bool? n = null;
bool? f = false;
__P((t & n).ToString());
__P((n & f).ToString());
__P((f & t).ToString());
__Check("\n\nFalse");
