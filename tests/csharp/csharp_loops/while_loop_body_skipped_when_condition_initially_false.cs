// vybe-test: csharp/csharp_loops/while_loop_body_skipped_when_condition_initially_false
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

int count=0; while(false) count++; Console.WriteLine(count);
