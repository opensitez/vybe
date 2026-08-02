// vybe-test: csharp/csharp_oop_inheritance/method_hiding_with_new_keyword_selects_by_static_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Parent { public string Name() => "Parent"; }
class Child : Parent { public new string Name() => "Child"; }
Parent p = new Child();
__Check((p.Name()).ToString(), "Parent");
