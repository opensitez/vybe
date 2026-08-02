// vybe-test: csharp/csharp_goto_switch_labels/switch_fallthrough_via_goto_case_only
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int v = 1;
int total = 0;
switch (v) {
    case 1: total += 10; goto case 2;
    case 2: total += 1; break;
}
__Check((total).ToString(), "11");
