// vybe-test: csharp/more_classes/record_with_body
// origin: languages/csharp/tests/csharp/test_more_classes.rs

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

record Car(string Make, int Year) {
            public string Info() {
                return Make + " " + Year;
            }
        }
        var c = new Car("Toyota", 2024);
        __P((c.Info()).ToString());
__Check("Toyota 2024");
