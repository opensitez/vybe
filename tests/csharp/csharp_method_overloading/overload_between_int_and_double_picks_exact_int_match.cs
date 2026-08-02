// vybe-test: csharp/csharp_method_overloading/overload_between_int_and_double_picks_exact_int_match
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Kind(int n)=>"int";
string Kind(double d)=>"double";
__Check((Kind(5)).ToString(), "int");
__Check((Kind(5.0)).ToString(), "double");
