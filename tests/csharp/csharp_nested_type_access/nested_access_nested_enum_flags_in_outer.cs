// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_flags_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Auth{[System.Flags] public enum Perm{None=0,Read=1,Write=2} public Perm All()=>Perm.Read|Perm.Write;} __Check(((int)new Auth().All()).ToString(), "3");
