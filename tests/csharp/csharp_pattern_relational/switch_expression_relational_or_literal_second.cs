// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_or_literal_second
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int code=500; __Check((code switch{200=>"ok",404 or 500=>"err",_=>"?"}).ToString(), "err");
