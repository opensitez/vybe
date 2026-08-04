// vybe-test: csharp/csharp_nullable_value_deep/nullable_bool_lifted_or_both_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

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

bool? a=null; bool? b=null; __P(((a|b).HasValue).ToString());
__Check("False");
