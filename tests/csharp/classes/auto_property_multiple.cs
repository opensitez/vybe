// vybe-test: csharp/classes/auto_property_multiple
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Car {
            public string Model { get; set; }
            public int Year { get; set; }
            public Car(string m, int y) { this.Model = m; this.Year = y; }
        }
        var c = new Car("Tesla", 2024);
        __Check((c.Model).ToString(), "Tesla");
        __Check((c.Year).ToString(), "2024");
