// vybe-test: csharp/csharp_pattern_property/switch_expression_property_when_guard_on_fields
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

class Pair { public int A; public int B; } string Sign(object o)=>o switch{Pair{A:var x,B:var y} when x==y=>"eq",Pair{A:var x,B:var y}=>"neq",_=>"?"}; __P((Sign(new Pair{A=3,B=3})).ToString());
__Check("eq");
