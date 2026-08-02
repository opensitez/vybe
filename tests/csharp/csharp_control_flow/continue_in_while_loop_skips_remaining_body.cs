// vybe-test: csharp/csharp_control_flow/continue_in_while_loop_skips_remaining_body
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int n = 0;
int sum = 0;
while (n < 5) {
    n++;
    if (n % 2 == 0) continue;
    sum += n;
}
Console.WriteLine(sum);
