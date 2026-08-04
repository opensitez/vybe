// vybe-test: csharp/csharp_static_type_behaviors/static_property_tracks_number_of_created_objects
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

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
    public static int Created { get; private set; }
    public Token() { Created++; }
}
new Token();
new Token();
new Token();
__P((Token.Created).ToString());
__Check("3");
