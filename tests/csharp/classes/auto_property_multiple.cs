// vybe-test: csharp/classes/auto_property_multiple
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var c = new Car("Tesla", 2024);
__P((c.Model).ToString());
__P((c.Year).ToString());
__Check("Tesla\n2024");

class Car {
            public string Model { get; set; }
            public int Year { get; set; }
            public Car(string m, int y) { this.Model = m; this.Year = y; }
        }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
