// vybe-test: csharp/csharp_oop_inheritance/is_operator_checks_runtime_type_in_hierarchy
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A { }
class B : A { }
object obj = new B();
__Check((obj is A).ToString(), "True");
__Check((obj is B).ToString(), "True");
