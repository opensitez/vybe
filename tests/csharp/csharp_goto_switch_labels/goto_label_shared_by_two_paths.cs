// vybe-test: csharp/csharp_goto_switch_labels/goto_label_shared_by_two_paths
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int flag = 1;
int result = 0;
if (flag == 0) goto finish;
result = 5;
finish:
__Check((result).ToString(), "5");
