// vybe-test: csharp/csharp_json_serialization/json_deserialize_dictionary_from_object_json
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=System.Text.Json.JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string,int>>("{"a":1,"b":2}");
__Check((d["a"]).ToString(), "1"); __Check((d.Count).ToString(), "2");
