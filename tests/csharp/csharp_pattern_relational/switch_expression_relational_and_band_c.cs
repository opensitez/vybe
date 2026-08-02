// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_and_band_c
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=55; __Check((n switch{>=90=>"A",>=70 and <90=>"B",>=50 and <70=>"C",_=>"F"}).ToString(), "C");
