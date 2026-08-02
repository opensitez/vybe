// vybe-test: csharp/csharp_nameof_expressions/nameof_overloaded_method_uses_simple_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Calc{public int Add(int a,int b)=>a+b; public double Add(double a,double b)=>a+b;} __Check((nameof(Calc.Add)).ToString(), "Add");
