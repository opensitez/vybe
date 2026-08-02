// vybe-test: csharp/csharp_checked_overflow_throws/checked_add_on_int_max_value_throws_overflow_exception
// origin: languages/csharp/tests/csharp/test_csharp_checked_overflow_throws.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string outcome = "ok";
try {
    checked {
        int value = int.MaxValue;
        value += 1;
    }
} catch (System.OverflowException) {
    outcome = "overflow";
}
__Check((outcome).ToString(), "overflow");
