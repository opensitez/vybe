// vybe-test: csharp/csharp_goto_switch_labels/goto_label_after_switch_accumulates
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int code = 2;
int acc = 0;
switch (code) {
    case 1: acc += 1; break;
    case 2: acc += 2; goto default;
    default: acc += 100; break;
}
__Check((acc).ToString(), "102");
