// vybe-test: csharp/csharp_json_serialization/json_options_case_insensitive_deserialization
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

class Item{public string Label{get;set;}}
var opts=new System.Text.Json.JsonSerializerOptions{PropertyNameCaseInsensitive=true};
var item=System.Text.Json.JsonSerializer.Deserialize<Item>("{"label":"x"}",opts);
__P((item.Label).ToString());
__Check("x");
