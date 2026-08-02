// vybe-test: csharp/csharp_ref_out_in/out_inline_declaration_in_method_call
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool ok = int.TryParse("42", out int result);
__Check((ok).ToString(), "True"); __Check((result).ToString(), "42");
