// vybe-test: csharp/csharp_fixed_unsafe_buffers/fixed_int_ptr_sum_via_loop_accumulates
// origin: languages/csharp/tests/csharp/test_csharp_fixed_unsafe_buffers.rs

int[] arr={1,2,3,4}; int sum=0; unsafe{fixed(int* ptr=&arr[0]){for(int i=0;i<4;i++){sum+=ptr[i];}}} Console.WriteLine(sum);
