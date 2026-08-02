// vybe-test: csharp/csharp_nameof_expressions/nameof_instance_method_on_type_returns_method_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Worker{public void Run(){}} __Check((nameof(Worker.Run)).ToString(), "Run");
