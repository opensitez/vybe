// vybe-test: csharp/csharp_goto_switch_labels/goto_label_switch_mix_with_loop_break
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            log += "0";
            break;
        case 1:
            log += "1";
            break;
    }
    if (i == 1) break;
}
Console.WriteLine(log);
