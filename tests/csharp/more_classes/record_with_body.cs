// vybe-test: csharp/more_classes/record_with_body
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var c = new Car("Toyota", 2024);
__P((c.Info()).ToString());
__Check("Toyota 2024");

record Car(string Make, int Year) {
            public string Info() {
                return Make + " " + Year;
            }
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
