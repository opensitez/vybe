// vybe-test: csharp/csharp_access_modifiers/internal_member_accessible_within_same_assembly
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Library{internal string Tag="v1";}
__Check((new Library().Tag).ToString(), "v1");
