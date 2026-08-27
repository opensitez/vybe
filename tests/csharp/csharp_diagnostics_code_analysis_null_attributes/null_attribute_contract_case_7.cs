// vybe-test: csharp/csharp_diagnostics_code_analysis_null_attributes/null_attribute_contract_case_7

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

bool TryGetItem([System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string res) {
    res = "Valid_7";
    return true;
}
bool ok = TryGetItem(out string val);
__P(ok.ToString());
__P(val);
__Check("True\nValid_7");
