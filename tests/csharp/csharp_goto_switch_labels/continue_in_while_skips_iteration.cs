// vybe-test: csharp/csharp_goto_switch_labels/continue_in_while_skips_iteration
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int n = 0;
int sum = 0;
while (n < 5) {
    n++;
    if (n == 3) continue;
    sum += n;
}
Console.WriteLine(sum);
