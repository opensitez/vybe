// vybe-test: csharp/csharp_oop_polymorphism/as_operator_returns_null_for_incompatible_cast
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{} class B{}
object o=new A();
__Check((o as B==null).ToString(), "True");
