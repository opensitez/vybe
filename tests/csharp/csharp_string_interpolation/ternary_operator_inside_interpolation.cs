// vybe-test: csharp/csharp_string_interpolation/ternary_operator_inside_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n=7; __Check(($"{(n%2==0?"even":"odd")}").ToString(), "odd");
