// vybe-test: csharp/csharp_nullable_value_type_arithmetic/nullable_addition_with_both_operands_present_yields_sum
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

int? left = 4;
int? right = 6;
int? sum = left + right;
__P((sum.HasValue).ToString());
__P((sum.Value).ToString());
__Check("True\n10");
