// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_string_with_conditional_expression
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

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

int n=4; string text=$"""{ (n%2==0 ? "even" : "odd") }"""; __P((text.Trim()).ToString());
__Check("even");
