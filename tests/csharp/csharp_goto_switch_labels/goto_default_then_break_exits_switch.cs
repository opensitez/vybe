// vybe-test: csharp/csharp_goto_switch_labels/goto_default_then_break_exits_switch
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int n = 0;
string r = "";
switch (n) {
    case 0:
        goto default;
    default:
        r = "done";
        break;
}
__Check((r).ToString(), "done");
