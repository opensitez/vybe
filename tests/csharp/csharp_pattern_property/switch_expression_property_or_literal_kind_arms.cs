// vybe-test: csharp/csharp_pattern_property/switch_expression_property_or_literal_kind_arms
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

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

class Msg { public string Kind; } string Label(object o)=>o switch{Msg{Kind:"err" or "fail"}=>"bad",_=>"ok"}; __P((Label(new Msg{Kind="fail"})).ToString());
__Check("bad");
