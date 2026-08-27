// vybe-test: csharp/csharp_json_custom_converter_polymorphic/json_options_case_18

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

var opts = new System.Text.Json.JsonSerializerOptions();
opts.PropertyNameCaseInsensitive = true;
__P(opts.PropertyNameCaseInsensitive.ToString());
__Check("True");
