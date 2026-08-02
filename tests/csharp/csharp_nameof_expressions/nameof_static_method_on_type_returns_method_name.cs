// vybe-test: csharp/csharp_nameof_expressions/nameof_static_method_on_type_returns_method_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class MathUtil{public static int Double(int n)=>n*2;} __Check((nameof(MathUtil.Double)).ToString(), "Double");
