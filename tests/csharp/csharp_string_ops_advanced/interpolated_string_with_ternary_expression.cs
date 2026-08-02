// vybe-test: csharp/csharp_string_ops_advanced/interpolated_string_with_ternary_expression
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=5; __Check(($"{(n>3?"big":"small")}").ToString(), "big");
