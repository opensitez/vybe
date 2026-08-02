// vybe-test: csharp/csharp_local_function_static/local_function_capture_params_array
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

int SumAll(int[] nums){int Total(int n){int s=0; for(int i=0;i<nums.Length;i++){s+=nums[i];} return s;} return Total(nums.Length);} Console.WriteLine(SumAll(new int[]{1,2,3}));
