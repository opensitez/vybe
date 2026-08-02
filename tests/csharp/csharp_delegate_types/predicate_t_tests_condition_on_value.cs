// vybe-test: csharp/csharp_delegate_types/predicate_t_tests_condition_on_value
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Predicate<string> isLong = s => s.Length > 4;
__Check((isLong("hello")).ToString(), "True");
__Check((isLong("hi")).ToString(), "False");
