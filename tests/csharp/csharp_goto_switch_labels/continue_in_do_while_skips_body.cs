// vybe-test: csharp/csharp_goto_switch_labels/continue_in_do_while_skips_body
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int n = 0;
int sum = 0;
do {
    n++;
    if (n == 2) continue;
    sum += n;
} while (n < 4);
Console.WriteLine(sum);
