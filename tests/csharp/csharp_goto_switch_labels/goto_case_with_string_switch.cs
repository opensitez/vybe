// vybe-test: csharp/csharp_goto_switch_labels/goto_case_with_string_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string key = "b";
string r = "";
switch (key) {
    case "a": r += "A"; goto case "b";
    case "b": r += "B"; break;
    case "c": r += "C"; break;
}
__Check((r).ToString(), "B");
