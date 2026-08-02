// vybe-test: csharp/csharp_nested_control_flow/nested_switch_break_exits_only_inner_switch
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

string report = "";
for (int i = 0; i < 2; i++) {
    switch (i) {
        case 0:
            switch (i) {
                case 0:
                    report += "inner;";
                    break;
            }
            report += "after-inner;";
            break;
        case 1:
            report += "tail;";
            break;
    }
}
Console.WriteLine(report);
