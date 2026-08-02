// vybe-test: csharp/csharp_goto_switch_labels/goto_case_on_zero_value
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int v = 0;
string r = "";
switch (v) {
    case 0: r += "0"; goto case 1;
    case 1: r += "1"; break;
}
__Check((r).ToString(), "01");
