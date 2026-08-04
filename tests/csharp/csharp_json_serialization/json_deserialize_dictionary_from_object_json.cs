// vybe-test: csharp/csharp_json_serialization/json_deserialize_dictionary_from_object_json
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

var d=System.Text.Json.JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string,int>>("{"a":1,"b":2}");
__P((d["a"]).ToString()); __P((d.Count).ToString());
__Check("1\n2");
