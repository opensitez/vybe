// vybe-test: csharp/csharp_json_serialization/json_roundtrip_preserves_nested_object
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

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

class Inner{public int X{get;set;}}
class Outer{public Inner Child{get;set;}}
var orig=new Outer{Child=new Inner{X=42}};
var json=System.Text.Json.JsonSerializer.Serialize(orig);
var back=System.Text.Json.JsonSerializer.Deserialize<Outer>(json);
__P((back.Child.X).ToString());
__Check("42");
