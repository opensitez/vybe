// vybe-test: csharp/csharp_oop_inheritance/protected_member_is_accessible_in_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { protected int Secret = 42; }
class Child : Base { public int Get() => Secret; }
__Check((new Child().Get()).ToString(), "42");
