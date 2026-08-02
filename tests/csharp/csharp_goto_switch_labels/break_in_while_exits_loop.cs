// vybe-test: csharp/csharp_goto_switch_labels/break_in_while_exits_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int n = 0;
while (true) {
    n++;
    if (n == 3) break;
}
Console.WriteLine(n);
