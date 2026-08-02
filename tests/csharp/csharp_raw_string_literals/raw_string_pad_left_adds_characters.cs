// vybe-test: csharp/csharp_raw_string_literals/raw_string_pad_left_adds_characters
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text="""7"""; __Check((text.PadLeft(3,'0')).ToString(), "007");
