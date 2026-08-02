// vybe-test: csharp/csharp_nested_control_flow/while_loop_with_continue_reaches_next_condition_check
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int n = 0;
int sum = 0;
while (n < 5) {
    n++;
    if (n == 3) continue;
    sum += n;
}
Console.WriteLine(sum);
