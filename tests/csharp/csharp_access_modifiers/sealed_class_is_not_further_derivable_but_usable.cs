// vybe-test: csharp/csharp_access_modifiers/sealed_class_is_not_further_derivable_but_usable
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

sealed class Final{public int Value=99;}
__Check((new Final().Value).ToString(), "99");
