// vybe-test: csharp/csharp_oop/record_with_body
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Product(string Name, double Price) {
    public string Display() { return Name + ": $" + Price; }
}
var p = new Product("Widget", 9.99);
__Check((p.Display()).ToString(), "Widget: $9.99");
