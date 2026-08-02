// vybe-test: csharp/csharp_strings/string_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string name = "Alice";
int age = 30;
__Check(($"{name} is {age}").ToString(), "Alice is 30");
