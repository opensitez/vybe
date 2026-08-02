// vybe-test: csharp/more_classes/lambda_as_callback_to_method
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Util {
            public int Apply(int x) {
                return x * 2;
            }
        }
        var u = new Util();
        __Check((u.Apply(21)).ToString(), "42");
