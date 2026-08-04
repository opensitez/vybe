// vybe-test: csharp/csharp_equality_contracts/boxed_value_types_with_same_numeric_value_compare_equal_with_equals
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

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

object left = 42;
object right = 42;
__P((left.Equals(right)).ToString());
__Check("True");
