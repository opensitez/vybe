// vybe-test: csharp/csharp_reflection_activation/gettype_on_instance_reports_runtime_type_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object value = "hello"; __Check((value.GetType().Name).ToString(), "String");
