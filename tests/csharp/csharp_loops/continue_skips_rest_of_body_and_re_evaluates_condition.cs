// vybe-test: csharp/csharp_loops/continue_skips_rest_of_body_and_re_evaluates_condition
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

int s=0; for(int i=0;i<5;i++) { if(i%2==0) continue; s+=i; } Console.WriteLine(s);
