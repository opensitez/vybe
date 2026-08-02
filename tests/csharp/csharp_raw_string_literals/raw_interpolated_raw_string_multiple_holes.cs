// vybe-test: csharp/csharp_raw_string_literals/raw_interpolated_raw_string_multiple_holes
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string name="Ada"; int age=36; string text=$"""{name} is {age}"""; __Check((text).ToString(), "Ada is 36");
