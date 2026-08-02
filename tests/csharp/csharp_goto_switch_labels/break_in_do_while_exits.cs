// vybe-test: csharp/csharp_goto_switch_labels/break_in_do_while_exits
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int n = 0;
do {
    n++;
    if (n == 2) break;
} while (n < 10);
Console.WriteLine(n);
