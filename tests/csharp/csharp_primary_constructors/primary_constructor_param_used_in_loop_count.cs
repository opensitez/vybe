// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_loop_count
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

class Repeat(int times) { public int Run() { int t = 0; for (int i = 0; i < times; i++) t++; return t; } }
Console.WriteLine(new Repeat(4).Run());
