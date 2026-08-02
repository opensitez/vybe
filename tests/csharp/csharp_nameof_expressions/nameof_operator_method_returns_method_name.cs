// vybe-test: csharp/csharp_nameof_expressions/nameof_operator_method_returns_method_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Vector{public static Vector operator +(Vector a,Vector b)=>a; public int X;} __Check((nameof(Vector.op_Addition)).ToString(), "op_Addition");
