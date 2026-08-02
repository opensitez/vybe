// vybe-test: csharp/csharp_reflection_activation/get_generic_type_definition_returns_open_generic_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> { } __Check((typeof(Box<int>).GetGenericTypeDefinition().Name.Contains("Box")).ToString(), "True");
