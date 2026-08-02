// vybe-test: csharp/csharp_nameof_expressions/nameof_nested_type_member_returns_member_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer{public class Inner{public int Value;}} __Check((nameof(Outer.Inner.Value)).ToString(), "Value");
