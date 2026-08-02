// vybe-test: csharp/csharp_nested_classes/nested_class_fully_qualified_via_outer_name
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Container{public class Item{public int Value=7;}}
var item=new Container.Item();
__Check((item.Value).ToString(), "7");
