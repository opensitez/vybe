// vybe-test: csharp/csharp_nested_control_flow/continue_inside_inner_loop_skips_remaining_body_but_not_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int sum = 0;
for (int outer = 0; outer < 2; outer++) {
    for (int inner = 0; inner < 3; inner++) {
        if (inner == 1) continue;
        sum += inner;
    }
}
Console.WriteLine(sum);
