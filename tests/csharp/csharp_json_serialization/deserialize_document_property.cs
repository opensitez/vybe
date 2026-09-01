// vybe-test: csharp/csharp_json_serialization/deserialize_document_property
// Expectations generated from .NET SDK 10 — never hand-written.

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using var d = System.Text.Json.JsonDocument.Parse(System.Text.Json.JsonSerializer.Serialize(new { id = 9 }));
__P(d.RootElement.GetProperty("id").GetInt32().ToString());
__Check("9");
