// vybe-test: csharp/csharp_nullable_semantics/nullable_bool_supports_three_state_logic
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool? a = true, b = null;
__Check((a == true).ToString(), "True");
__Check((b == null).ToString(), "True");
