// vybe-test: csharp/csharp_local_function_static/local_function_while_loop_with_capture
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

int Count(int n){int i=0; int acc=0; while(i<n){int Step(int x)=>acc+x; acc=Step(i+1); i++;} return acc;} Console.WriteLine(Count(3));
