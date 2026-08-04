// vybe-test: csharp/csharp_pattern_property/switch_expression_property_default_after_specific_arms
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

class Token { public string Kind; } string Name(object o)=>o switch{Token{Kind:"add"}=>"plus",Token{Kind:"sub"}=>"minus",_=>"other"}; __P((Name(new Token{Kind="mul"})).ToString());
__Check("other");
