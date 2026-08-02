// vybe-test: csharp/csharp_reflection/typeof_on_string_has_correct_full_name
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((typeof(string).FullName).ToString(), "System.String");
