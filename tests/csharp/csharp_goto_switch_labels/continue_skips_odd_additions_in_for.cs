// vybe-test: csharp/csharp_goto_switch_labels/continue_skips_odd_additions_in_for
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int sum = 0;
for (int i = 1; i <= 6; i++) {
    if (i % 2 == 0) continue;
    sum += i;
}
Console.WriteLine(sum);
