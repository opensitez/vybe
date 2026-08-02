// vybe-test: csharp/csharp_json_serialization/serialize_simple_object_to_json_string
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var obj=new{Name="Alice",Age=30};
string json=System.Text.Json.JsonSerializer.Serialize(obj);
__Check((json.Contains("Alice")).ToString(), "True");
