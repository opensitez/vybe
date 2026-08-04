// vybe-test: csharp/csharp_json_serialization/serialize_simple_object_to_json_string
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

var obj=new{Name="Alice",Age=30};
string json=System.Text.Json.JsonSerializer.Serialize(obj);
__P((json.Contains("Alice")).ToString());
__Check("True");
