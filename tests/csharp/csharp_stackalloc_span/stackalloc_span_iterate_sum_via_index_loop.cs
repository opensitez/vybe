// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_iterate_sum_via_index_loop
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

System.Span<int> span=stackalloc int[3]{1,2,3}; int sum=0; for(int i=0;i<span.Length;i++){sum+=span[i];} Console.WriteLine(sum);
