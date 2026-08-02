// vybe-test: csharp/csharp_reflection_activation/enum_type_reports_isenum_true
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State { Ready } __Check((typeof(State).IsEnum).ToString(), "True");
