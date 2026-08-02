// vybe-test: csharp/csharp_goto_switch_labels/continue_outer_not_available_only_inner
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int count = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 2; j++) {
        if (j == 0) continue;
        count++;
    }
}
Console.WriteLine(count);
