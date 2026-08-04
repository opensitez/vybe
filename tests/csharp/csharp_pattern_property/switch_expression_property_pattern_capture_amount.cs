// vybe-test: csharp/csharp_pattern_property/switch_expression_property_pattern_capture_amount
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

class Wallet { public int Balance; } int Read(object o)=>o switch{Wallet{Balance:var b}=>b,_=>-1}; __P((Read(new Wallet{Balance=42})).ToString());
__Check("42");
