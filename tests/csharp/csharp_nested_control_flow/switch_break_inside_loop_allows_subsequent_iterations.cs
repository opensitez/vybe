// vybe-test: csharp/csharp_nested_control_flow/switch_break_inside_loop_allows_subsequent_iterations
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int sum = 0;
for (int i = 0; i < 4; i++) {
    switch (i) {
        case 1:
        case 2:
            sum += 10;
            break;
        default:
            sum += 1;
            break;
    }
}
Console.WriteLine(sum);
