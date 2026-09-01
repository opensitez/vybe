// vybe-test: csharp/csharp_json_type_info_contract_modifiers/typeinfo_doc_int
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

__P((System.Text.Json.JsonDocument.Parse("{\"a\":2}").RootElement.GetProperty("a").GetInt32()).ToString());
__Check("2");
