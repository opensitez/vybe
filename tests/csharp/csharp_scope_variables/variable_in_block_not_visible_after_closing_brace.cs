// vybe-test: csharp/csharp_scope_variables/variable_in_block_not_visible_after_closing_brace
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int outer = 1;
{
    int inner = 2;
    outer = inner;
}
__Check((outer).ToString(), "2");
