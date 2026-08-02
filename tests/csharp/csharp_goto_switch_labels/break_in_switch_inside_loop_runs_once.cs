// vybe-test: csharp/csharp_goto_switch_labels/break_in_switch_inside_loop_runs_once
// origin: languages/csharp/tests/csharp/test_csharp_goto_switch_labels.rs

string report = "";
for (int i = 0; i < 3; i++) {
    switch (i) {
        case 0: report += "a"; break;
        case 1: report += "b"; break;
        case 2: report += "c"; break;
    }
}
Console.WriteLine(report);
