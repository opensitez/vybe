// vybe-test: csharp/csharp_enum_flags_operations/enum_to_string_returns_declared_identifier
// origin: languages/csharp/tests/csharp/test_csharp_enum_flags_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Status { Idle, Running, Done }
__Check((Status.Running.ToString()).ToString(), "Running");
