// vybe-test: csharp/csharp_pattern_property/switch_expression_property_when_false_falls_through
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

class Item { public int Q; } string Flag(object o)=>o switch{Item{Q:var q} when q>10=>"big",Item{Q:var q}=>"small",_=>"?"}; __P((Flag(new Item{Q=3})).ToString());
__Check("small");
