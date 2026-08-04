// vybe-test: csharp/csharp_anonymous_type_inference/anonymous_type_infers_property_names_from_simple_identifier_expressions
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_type_inference.rs

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

int width = 4;
string label = "box";
var shape = new { width, label };
__P((shape.width).ToString());
__P((shape.label).ToString());
__Check("4\nbox");
