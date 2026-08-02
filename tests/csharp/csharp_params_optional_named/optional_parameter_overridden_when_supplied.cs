// vybe-test: csharp/csharp_params_optional_named/optional_parameter_overridden_when_supplied
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Greet(string name, string prefix="Hello") => prefix+" "+name;
__Check((Greet("World","Hi")).ToString(), "Hi World");
