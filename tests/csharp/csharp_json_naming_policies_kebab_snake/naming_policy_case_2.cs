// vybe-test: csharp/csharp_json_naming_policies_kebab_snake/naming_policy_case_2

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

string converted = System.Text.Json.JsonNamingPolicy.SnakeCaseLower.ConvertName("ItemValue2");
__P((converted.StartsWith("item_value")).ToString());
__Check("True");
