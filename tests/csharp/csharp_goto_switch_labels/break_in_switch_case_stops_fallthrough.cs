// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_case_stops_fallthrough
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 1;
string r = "";
switch (n) {
    case 1: r += "x"; break;
    case 2: r += "y"; break;
}
__Check((r).ToString(), "x");
