// vybe-test: csharp/csharp_json_serialization/json_roundtrip_preserves_nested_object
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Inner{public int X{get;set;}}
class Outer{public Inner Child{get;set;}}
var orig=new Outer{Child=new Inner{X=42}};
var json=System.Text.Json.JsonSerializer.Serialize(orig);
var back=System.Text.Json.JsonSerializer.Deserialize<Outer>(json);
__Check((back.Child.X).ToString(), "42");
