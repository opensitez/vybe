// vybe-test: csharp/csharp_params_optional_named/optional_with_null_default_allows_omission
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Label(string text, string tag=null) => tag==null?text:$"[{tag}]{text}";
__Check((Label("msg")).ToString(), "msg");
__Check((Label("msg","info")).ToString(), "[info]msg");
