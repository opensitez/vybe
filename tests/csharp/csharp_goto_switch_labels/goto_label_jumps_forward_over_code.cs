// vybe-test: csharp/csharp_goto_switch_labels/goto_label_jumps_forward_over_code
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 0;
start:
x++;
if (x < 3) goto start;
__Check((x).ToString(), "3");
