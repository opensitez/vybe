// vybe-test: csharp/csharp_json_serialization/json_serialize_list_produces_array_syntax
// origin: languages/csharp/tests/csharp/test_csharp_json_serialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list=new System.Collections.Generic.List<int>{1,2,3};
string json=System.Text.Json.JsonSerializer.Serialize(list);
__Check((json).ToString(), "[1,2,3]");
