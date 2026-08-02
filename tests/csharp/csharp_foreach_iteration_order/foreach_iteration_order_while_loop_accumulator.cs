// vybe-test: csharp/csharp_foreach_iteration_order/foreach_iteration_order_while_loop_accumulator
// origin: languages/csharp/tests/csharp/test_csharp_foreach_iteration_order.rs

// foreach_iteration_order
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);
