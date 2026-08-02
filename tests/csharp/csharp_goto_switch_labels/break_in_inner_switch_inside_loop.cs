// vybe-test: csharp/csharp_goto_switch_labels/break_in_inner_switch_inside_loop
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

string log = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            switch (i) {
                case 0: log += "in"; break;
            }
            log += ";";
            break;
        case 1: log += "out"; break;
    }
}
Console.WriteLine(log);
