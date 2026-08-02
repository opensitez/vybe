// vybe-test: csharp/csharp_goto_switch_labels/labeled_continue_in_for_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int sum = 0;
for (int i = 0; i < 5; i++) {
    if (i == 2) continue;
    sum += i;
}
Console.WriteLine(sum);
