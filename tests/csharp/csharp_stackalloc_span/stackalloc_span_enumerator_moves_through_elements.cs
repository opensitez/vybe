// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_enumerator_moves_through_elements
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

System.Span<int> span=stackalloc int[3]{2,4,6}; int sum=0; foreach(int v in span){sum+=v;} Console.WriteLine(sum);
