// vybe-test: csharp/csharp_goto_switch_labels/goto_default_from_case_arm
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
    case 1:
        r += "start";
        goto default;
    default:
        r += ":default";
        break;
}
__Check((r).ToString(), "start:default");
