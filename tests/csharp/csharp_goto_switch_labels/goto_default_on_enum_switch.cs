// vybe-test: csharp/csharp_goto_switch_labels/goto_default_on_enum_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green }
Color c = (Color)9;
string name = "";
switch (c) {
    case Color.Red: name = "R"; break;
    case Color.Green: name = "G"; break;
    default: name = "?"; break;
}
__Check((name).ToString(), "?");
