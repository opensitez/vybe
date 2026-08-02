// vybe-test: csharp/csharp_loops/do_while_body_runs_at_least_once_when_condition_false
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

int count=0; do { count++; } while(false); Console.WriteLine(count);
