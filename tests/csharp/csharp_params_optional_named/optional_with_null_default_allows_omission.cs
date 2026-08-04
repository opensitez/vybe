// vybe-test: csharp/csharp_params_optional_named/optional_with_null_default_allows_omission
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

string Label(string text, string tag=null) => tag==null?text:$"[{tag}]{text}";
__P((Label("msg")).ToString());
__P((Label("msg","info")).ToString());
__Check("msg\n[info]msg");
