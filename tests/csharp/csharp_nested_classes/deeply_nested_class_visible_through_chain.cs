// vybe-test: csharp/csharp_nested_classes/deeply_nested_class_visible_through_chain
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class A{public class B{public class C{public int V=3;}}}
__Check((new A.B.C().V).ToString(), "3");
