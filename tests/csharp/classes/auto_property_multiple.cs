// vybe-test: csharp/classes/auto_property_multiple
// origin: languages/csharp/tests/csharp/test_classes.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Car {
            public string Model { get; set; }
            public int Year { get; set; }
            public Car(string m, int y) { this.Model = m; this.Year = y; }
        }
        var c = new Car("Tesla", 2024);
        __P((c.Model).ToString());
        __P((c.Year).ToString());
__Check("Tesla\n2024");
