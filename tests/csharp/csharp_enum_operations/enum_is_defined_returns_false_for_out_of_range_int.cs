// vybe-test: csharp/csharp_enum_operations/enum_is_defined_returns_false_for_out_of_range_int
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Level{Low=0,Mid=1,High=2}
__Check((System.Enum.IsDefined(typeof(Level), 99)).ToString(), "False");
