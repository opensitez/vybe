// vybe-test: csharp/csharp_ref_out_in/in_parameter_prevents_copy_and_is_readonly
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Sum3(in int a, in int b, in int c) => a+b+c;
__Check((Sum3(1,2,3)).ToString(), "6");
