// vybe-test: csharp/csharp_json_serialization/deserialize_json_string_to_typed_record
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

record Person(string Name,int Age);
string json="{"Name":"Bob","Age":25}";
var p=System.Text.Json.JsonSerializer.Deserialize<Person>(json);
__P((p.Name).ToString()); __P((p.Age).ToString());
__Check("Bob\n25");
