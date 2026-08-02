// vybe-test: csharp/csharp_exception_types/key_not_found_exception_thrown_by_dictionary_missing_key
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string result = "";
var map = new System.Collections.Generic.Dictionary<string,int>();
try { int v = map["nope"]; }
catch(System.Collections.Generic.KeyNotFoundException) { result = "missing"; }
__Check((result).ToString(), "missing");
