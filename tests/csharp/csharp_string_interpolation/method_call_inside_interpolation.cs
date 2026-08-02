// vybe-test: csharp/csharp_string_interpolation/method_call_inside_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="hello"; __Check(($"{s.ToUpper()}").ToString(), "HELLO");
