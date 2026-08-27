// vybe-test: csharp/csharp_local_function_static/local_function_capture_params_array
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

using static __Harness;

int SumAll(int[] nums){int Total(int n){int s=0; for(int i=0;i<nums.Length;i++){s+=nums[i];} return s;} return Total(nums.Length);}
__P((SumAll(new int[]{1,2,3})).ToString());
__Check("6");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
