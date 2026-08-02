// vybe-test: csharp/csharp_string_interpolation/expression_evaluated_inside_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a=3,b=4; __Check(($"{a}+{b}={a+b}").ToString(), "3+4=7");
