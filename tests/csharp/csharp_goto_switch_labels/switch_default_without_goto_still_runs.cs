// vybe-test: csharp/csharp_goto_switch_labels/switch_default_without_goto_still_runs
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int v = 5;
string tag = "";
switch (v) {
    case 1: tag = "one"; break;
    default: tag = "many"; break;
}
__Check((tag).ToString(), "many");
