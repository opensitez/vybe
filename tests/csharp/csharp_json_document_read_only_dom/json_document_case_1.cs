// vybe-test: csharp/csharp_json_document_read_only_dom/json_document_case_1

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

using var doc = System.Text.Json.JsonDocument.Parse("{\"count\":1}");
int count = doc.RootElement.GetProperty("count").GetInt32();
__P(count.ToString());
__Check("1");
