// vybe-test: csharp/more_classes/record_with_body
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Car(string Make, int Year) {
            public string Info() {
                return Make + " " + Year;
            }
        }
        var c = new Car("Toyota", 2024);
        __Check((c.Info()).ToString(), "Toyota 2024");
