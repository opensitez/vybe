// vybe-test: csharp/csharp_goto_switch_labels/break_vs_continue_in_same_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

string r = "";
for (int i = 0; i < 4; i++) {
    if (i == 1) continue;
    if (i == 3) break;
    r += i;
}
Console.WriteLine(r);
