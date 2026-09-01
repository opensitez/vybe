// vybe-test: csharp/csharp_json_document_read_only_dom/doc_enumerate_object_count
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

using var doc = System.Text.Json.JsonDocument.Parse("{\"count\":7,\"name\":\"ada\",\"ok\":true,\"ratio\":1.5,\"nil\":null,\"tags\":[10,20,30],\"inner\":{\"deep\":\"yes\"}}");
int n = 0;
foreach (var p in doc.RootElement.EnumerateObject()) n++;
__P(n.ToString());
__Check("7");
