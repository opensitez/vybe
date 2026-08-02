// vybe-test: csharp/csharp_goto_switch_labels/break_exits_inner_for_only
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int total = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 3; j++) {
        if (j == 1) break;
        total++;
    }
}
Console.WriteLine(total);
