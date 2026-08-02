// vybe-test: csharp/csharp_constructor_patterns/base_constructor_called_before_derived_body
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{public int Order;public A(){Order=1;}}
class B:A{public int Extra;public B():base(){Extra=2;}}
var b=new B();
__Check((b.Order).ToString(), "1"); __Check((b.Extra).ToString(), "2");
