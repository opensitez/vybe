// vybe-test: csharp/csharp_class_features/object_to_string_default_returns_type_name
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget{}
__Check((new Widget().ToString()).ToString(), "Widget");
