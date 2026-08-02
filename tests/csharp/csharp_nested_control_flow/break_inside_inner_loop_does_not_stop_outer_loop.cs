// vybe-test: csharp/csharp_nested_control_flow/break_inside_inner_loop_does_not_stop_outer_loop
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int total = 0;
for (int row = 0; row < 2; row++) {
    for (int col = 0; col < 4; col++) {
        if (col == 2) break;
        total += 1;
    }
}
Console.WriteLine(total);
