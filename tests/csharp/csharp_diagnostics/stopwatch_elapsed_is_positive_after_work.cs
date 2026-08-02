// vybe-test: csharp/csharp_diagnostics/stopwatch_elapsed_is_positive_after_work
// origin: languages/csharp/tests/csharp/test_csharp_diagnostics.rs

var sw=System.Diagnostics.Stopwatch.StartNew();
int s=0; for(int i=0;i<10000;i++) s+=i;
sw.Stop();
Console.WriteLine(sw.ElapsedMilliseconds>=0);
