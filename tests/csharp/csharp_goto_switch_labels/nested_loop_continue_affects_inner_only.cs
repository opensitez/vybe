// vybe-test: csharp/csharp_goto_switch_labels/nested_loop_continue_affects_inner_only
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int hits = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 3; j++) {
        if (j == 1) continue;
        hits++;
    }
}
Console.WriteLine(hits);
