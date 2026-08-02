// vybe-test: csharp/csharp_record_struct_deep/record_struct_foreach_sum
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

record struct V(int N); var sum=0; foreach(var v in new[]{new V(1),new V(2),new V(3)}) sum+=v.N; Console.WriteLine(sum);
