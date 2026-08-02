// vybe-test: csharp/classes/interface_basic
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IGreeter {
            string Greet();
        }
        class HelloGreeter : IGreeter {
            public string Greet() {
                return "Hello from interface!";
            }
        }
        var g = new HelloGreeter();
        __Check((g.Greet()).ToString(), "Hello from interface!");
