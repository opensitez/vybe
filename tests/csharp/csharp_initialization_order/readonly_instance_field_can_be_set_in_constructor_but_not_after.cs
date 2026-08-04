// vybe-test: csharp/csharp_initialization_order/readonly_instance_field_can_be_set_in_constructor_but_not_after
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

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

class Token {
    public readonly string Value;
    public Token(string value) { Value = value; }
}
var token = new Token("abc");
__P((token.Value).ToString());
__Check("abc");
