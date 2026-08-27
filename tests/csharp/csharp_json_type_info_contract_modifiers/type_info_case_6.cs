// vybe-test: csharp/csharp_json_type_info_contract_modifiers/type_info_case_6

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

var resolver = new System.Text.Json.Serialization.Metadata.DefaultJsonTypeInfoResolver();
__P((resolver != null).ToString());
__Check("True");
