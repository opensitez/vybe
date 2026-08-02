// vybe-test: csharp/csharp_goto_switch_labels/while_continue_restarts_without_increment_bug
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int i = 0;
int sum = 0;
while (i < 4) {
    i++;
    if (i == 2) continue;
    sum += i;
}
Console.WriteLine(sum);
