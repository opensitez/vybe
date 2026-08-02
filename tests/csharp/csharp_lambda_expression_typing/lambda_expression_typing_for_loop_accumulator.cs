// vybe-test: csharp/csharp_lambda_expression_typing/lambda_expression_typing_for_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expression_typing.rs

// lambda_expression_typing
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);
