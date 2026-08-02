// vybe-test: csharp/csharp_nameof_expressions/nameof_static_field_returns_field_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter{public static int Total=0;} __Check((nameof(Counter.Total)).ToString(), "Total");
