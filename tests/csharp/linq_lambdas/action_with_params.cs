// vybe-test: csharp/linq_lambdas/action_with_params
// origin: languages/csharp/tests/csharp/test_linq_lambdas.rs

using static __Harness;

Action<string, int> describe = (name, age) => {
    __P((name + " is " + age).ToString());
}
;
describe("Alice", 30);
describe("Bob", 25);
__Check("Alice is 30\nBob is 25");

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
