// vybe-test: csharp/csharp_throw_in_null_coalescing/null_coalescing_throw_expression_skipped_when_left_has_value
// origin: languages/csharp/tests/csharp/test_csharp_throw_in_null_coalescing.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string? present = "ok";
string value = present ?? throw new System.Exception("fail");
__Check((value).ToString(), "ok");
