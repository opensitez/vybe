// vybe-test: csharp/csharp_text_string_syntax_attributes/string_syntax_case_6

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

string jsonSyntax = System.Diagnostics.CodeAnalysis.StringSyntaxAttribute.Json;
__P(jsonSyntax);
__Check("Json");
