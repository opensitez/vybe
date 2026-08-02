// vybe-test: csharp/csharp_nameof_expressions/nameof_type_parameter_name_on_generic_method
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Factory{public T Build<T>(T value)=>value;} __Check((nameof(Factory.Build)).ToString(), "Build");
