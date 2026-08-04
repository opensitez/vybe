// vybe-test: csharp/csharp_json_serialization/json_serialize_list_produces_array_syntax
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

var list=new System.Collections.Generic.List<int>{1,2,3};
string json=System.Text.Json.JsonSerializer.Serialize(list);
__P((json).ToString());
__Check("[1,2,3]");
