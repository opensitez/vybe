// vybe-test: csharp/csharp_params_optional_named/named_argument_can_be_passed_out_of_order
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Concat(string a, string b, string c) => a+b+c;
__Check((Concat(c:"3",a:"1",b:"2")).ToString(), "123");
