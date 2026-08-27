// vybe-test: csharp/csharp_random_random/random_next_within_exclusive_upper_bound
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

using static __Harness;

var rng=new System.Random(42);
for(int i=0;i<100;i++){
    int v=rng.Next(10);
    if(v<0||v>=10){__P(("fail").ToString());return;}
}
__P(("pass").ToString());
__Check("pass");

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
