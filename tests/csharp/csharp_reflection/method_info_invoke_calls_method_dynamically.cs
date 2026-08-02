// vybe-test: csharp/csharp_reflection/method_info_invoke_calls_method_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Calc { public int Double(int n) => n * 2; }
var obj = new Calc();
var method = typeof(Calc).GetMethod("Double");
__Check((method.Invoke(obj, new object[]{5})).ToString(), "10");
