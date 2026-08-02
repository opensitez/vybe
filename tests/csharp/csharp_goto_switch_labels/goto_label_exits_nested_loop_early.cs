// vybe-test: csharp/csharp_goto_switch_labels/goto_label_exits_nested_loop_early
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int count = 0;
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 3; j++) {
        if (i == 1 && j == 1) goto finished;
        count++;
    }
}
finished:
Console.WriteLine(count);
