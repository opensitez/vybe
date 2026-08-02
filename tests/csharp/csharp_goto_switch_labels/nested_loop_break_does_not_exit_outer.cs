// vybe-test: csharp/csharp_goto_switch_labels/nested_loop_break_does_not_exit_outer
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int outerRuns = 0;
for (int i = 0; i < 2; i++) {
    for (int j = 0; j < 2; j++) {
        if (j == 1) break;
        outerRuns++;
    }
}
Console.WriteLine(outerRuns);
