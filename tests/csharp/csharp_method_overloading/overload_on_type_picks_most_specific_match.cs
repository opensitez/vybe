// vybe-test: csharp/csharp_method_overloading/overload_on_type_picks_most_specific_match
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Label(object o)=>"object";
string Label(string s)=>"string";
__Check((Label("hi")).ToString(), "string");
__Check((Label((object)"hi")).ToString(), "object");
