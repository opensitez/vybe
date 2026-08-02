// vybe-test: csharp/csharp_json_serialization/deserialize_json_string_to_typed_record
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Person(string Name,int Age);
string json="{"Name":"Bob","Age":25}";
var p=System.Text.Json.JsonSerializer.Deserialize<Person>(json);
__Check((p.Name).ToString(), "Bob"); __Check((p.Age).ToString(), "25");
