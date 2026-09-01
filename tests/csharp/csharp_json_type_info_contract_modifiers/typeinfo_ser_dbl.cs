// vybe-test: csharp/csharp_json_type_info_contract_modifiers/typeinfo_ser_dbl
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

__P((System.Text.Json.JsonSerializer.Serialize(new { d = 2.5 })).ToString());
__Check("{\"d\":2.5}");
