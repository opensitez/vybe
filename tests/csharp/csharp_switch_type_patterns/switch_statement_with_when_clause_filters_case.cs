// vybe-test: csharp/csharp_switch_type_patterns/switch_statement_with_when_clause_filters_case
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 8;
string size = n switch {
    < 0 => "neg",
    >= 0 and < 10 => "small",
    _ => "big"
};
__Check((size).ToString(), "small");
