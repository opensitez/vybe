// vybe-test: csharp/csharp_nameof_expressions/nameof_qualified_alias_target_type
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Text=System.String; __Check((nameof(Text)).ToString(), "Text");
