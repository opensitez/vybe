// vybe-test: csharp/common_patterns/enum_with_values
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Status { Active = 1, Inactive = 0, Pending = 2 }
__Check(((int)Status.Active).ToString(), "1");
__Check(((int)Status.Inactive).ToString(), "0");
__Check(((int)Status.Pending).ToString(), "2");
