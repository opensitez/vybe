// vybe-test: csharp/csharp_reflection_activation/get_generic_arguments_reports_type_argument_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> { } __Check((typeof(Box<int>).GetGenericArguments()[0].Name).ToString(), "Int32");
