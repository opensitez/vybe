// vybe-test: csharp/csharp_goto_switch_labels/goto_case_on_enum_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green, Blue }
Color c = Color.Red;
string name = "";
switch (c) {
    case Color.Red: name += "R"; goto case Color.Green;
    case Color.Green: name += "G"; break;
    case Color.Blue: name += "B"; break;
}
__Check((name).ToString(), "RG");
