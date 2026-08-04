// vybe-test: csharp/csharp_pattern_property/switch_expression_property_nested_capture_sum
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

class Inner { public int A; public int B; } class Wrap { public Inner Data; } int Sum(object o)=>o switch{Wrap{Data:{A:var a,B:var b}}=>a+b,_=>0}; __P((Sum(new Wrap{Data=new Inner{A=6,B=7}})).ToString());
__Check("13");
