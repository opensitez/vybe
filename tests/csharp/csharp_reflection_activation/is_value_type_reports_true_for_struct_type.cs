// vybe-test: csharp/csharp_reflection_activation/is_value_type_reports_true_for_struct_type
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((typeof(System.DateTime).IsValueType).ToString(), "True");
