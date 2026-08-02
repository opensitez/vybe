// vybe-test: csharp/csharp_method_overloading/generic_overload_less_specific_than_typed_overload
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Foo<T>(T v)=>"generic";
string Foo(int v)=>"specific";
__Check((Foo(1)).ToString(), "specific");
__Check((Foo("x")).ToString(), "generic");
