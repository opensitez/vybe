// vybe-test: csharp/csharp_json_serialization/json_options_case_insensitive_deserialization
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item{public string Label{get;set;}}
var opts=new System.Text.Json.JsonSerializerOptions{PropertyNameCaseInsensitive=true};
var item=System.Text.Json.JsonSerializer.Deserialize<Item>("{"label":"x"}",opts);
__Check((item.Label).ToString(), "x");
