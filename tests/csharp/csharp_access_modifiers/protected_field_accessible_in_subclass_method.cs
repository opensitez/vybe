// vybe-test: csharp/csharp_access_modifiers/protected_field_accessible_in_subclass_method
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{protected int Value=7;}
class B:A{public int Read()=>Value;}
__Check((new B().Read()).ToString(), "7");
