// vybe-test: csharp/csharp_nested_control_flow/do_while_executes_body_before_first_condition_test
// origin: languages/csharp/tests/csharp/test_csharp_nested_control_flow.rs

int count = 0;
do {
    count++;
} while (count < 1);
Console.WriteLine(count);
