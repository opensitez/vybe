// vybe-test: csharp/csharp_local_function_static/local_function_in_loop_accumulates
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

int Sum(int n){int total=0; for(int i=1;i<=n;i++){int Add(int x)=>total+x; total=Add(i);} return total;} Console.WriteLine(Sum(3));
