// vybe-test: csharp/csharp_control_flow/if_elseif_chain
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

int score = 75;
if (score >= 90) Console.WriteLine("A");
else if (score >= 80) Console.WriteLine("B");
else if (score >= 70) Console.WriteLine("C");
else Console.WriteLine("F");
