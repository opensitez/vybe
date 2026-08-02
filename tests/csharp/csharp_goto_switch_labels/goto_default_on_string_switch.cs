// vybe-test: csharp/csharp_goto_switch_labels/goto_default_on_string_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string key = "z";
string r = "";
switch (key) {
    case "a": r = "A"; break;
    case "b": r = "B"; break;
    default: r = "?"; break;
}
__Check((r).ToString(), "?");
