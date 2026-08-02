// vybe-test: csharp/csharp_nameof_expressions/nameof_extension_method_target_type_member
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Extensions{public static int Twice(this int n)=>n*2;} __Check((nameof(Extensions.Twice)).ToString(), "Twice");
