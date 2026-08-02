// vybe-test: csharp/csharp_goto_switch_labels/double_nested_loop_goto_label_escape
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int ticks = 0;
for (int a = 0; a < 2; a++) {
    for (int b = 0; b < 2; b++) {
        ticks++;
        if (ticks == 3) goto done;
    }
}
done:
Console.WriteLine(ticks);
