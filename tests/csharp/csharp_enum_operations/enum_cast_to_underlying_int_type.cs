// vybe-test: csharp/csharp_enum_operations/enum_cast_to_underlying_int_type
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Priority{Low=1,Medium=2,High=3} __Check(((int)Priority.High).ToString(), "3");
