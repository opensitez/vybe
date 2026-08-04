// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_setter_expression_updates_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

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

class Logger { public string last = ""; public string Last { get => last; set => last = value; } }
var l = new Logger(); l.Last = "ok"; __P((l.Last).ToString());
__Check("ok");
