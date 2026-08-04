// vybe-test: csharp/csharp_nullable_semantics/arithmetic_on_two_nullable_ints_with_values_produces_nullable_result
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

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

int? a=3, b=4; int? c=a+b; __P((c).ToString());
__Check("7");
