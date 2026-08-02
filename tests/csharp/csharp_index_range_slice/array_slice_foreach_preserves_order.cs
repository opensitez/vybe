// vybe-test: csharp/csharp_index_range_slice/array_slice_foreach_preserves_order
// origin: languages/csharp/tests/csharp/test_csharp_index_range_slice.rs

int[] data={1,2,3,4}; var slice=data[1..3]; int sum=0; foreach(var n in slice) sum+=n; Console.WriteLine(sum);
