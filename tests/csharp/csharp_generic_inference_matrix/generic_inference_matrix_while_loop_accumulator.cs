// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

// generic_inference_matrix
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
