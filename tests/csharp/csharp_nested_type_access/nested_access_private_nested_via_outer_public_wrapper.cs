// vybe-test: csharp/csharp_nested_type_access/nested_access_private_nested_via_outer_public_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Vault{class Key{public int Code=99;} public int Open()=>new Key().Code;} __Check((new Vault().Open()).ToString(), "99");
