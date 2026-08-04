// vybe-test: csharp/csharp_anonymous_types/anonymous_type_to_string_shows_property_values
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

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

var a=new{X=3,Y=4};
__P((a.ToString().Contains("X = 3")).ToString());
__Check("True");
