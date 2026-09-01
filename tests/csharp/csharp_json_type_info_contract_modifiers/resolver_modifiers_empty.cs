// vybe-test: csharp/csharp_json_type_info_contract_modifiers/resolver_modifiers_empty
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

var r = new System.Text.Json.Serialization.Metadata.DefaultJsonTypeInfoResolver();
__P(r.Modifiers.Count.ToString());
__Check("0");
