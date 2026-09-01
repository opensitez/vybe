// vybe-test: csharp/csharp_json_serialization/roundtrip_document_to_string
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

using var d = System.Text.Json.JsonDocument.Parse("{\"k\":\"v\"}");
__P(d.RootElement.GetProperty("k").GetString());
__Check("v");
