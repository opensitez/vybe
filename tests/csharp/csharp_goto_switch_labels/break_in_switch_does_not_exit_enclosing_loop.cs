// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_does_not_exit_enclosing_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

int i = 0;
while (i < 2) {
    switch (i) {
        case 0: i++; break;
        case 1: i++; break;
    }
}
Console.WriteLine(i);
