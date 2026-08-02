// vybe-test: csharp/csharp_access_modifiers/private_field_only_accessible_within_declaring_class
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Safe{private int secret=42; public int Get()=>secret;}
__Check((new Safe().Get()).ToString(), "42");
