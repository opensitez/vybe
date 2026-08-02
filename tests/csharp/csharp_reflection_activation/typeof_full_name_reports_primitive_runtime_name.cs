// vybe-test: csharp/csharp_reflection_activation/typeof_full_name_reports_primitive_runtime_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((typeof(int).FullName).ToString(), "System.Int32");
