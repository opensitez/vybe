// vybe-test: csharp/csharp_goto_switch_labels/switch_goto_default_from_middle_case
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 2;
string r = "";
switch (n) {
    case 1: r += "1"; break;
    case 2: r += "2"; goto default;
    default: r += "D"; break;
}
__Check((r).ToString(), "2D");
