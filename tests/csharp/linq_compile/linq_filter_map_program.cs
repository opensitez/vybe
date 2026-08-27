// vybe-test: csharp/linq_compile/linq_filter_map_program
// origin: languages/csharp/tests/csharp/test_linq_compile.rs
// vybe-test-mode: compile

using static __Harness;

var numbers = new List<int>();
numbers.Add(1);
numbers.Add(2);
numbers.Add(3);
numbers.Add(4);
numbers.Add(5);
numbers.Add(6);
numbers.Add(7);
numbers.Add(8);
var evens = numbers.Where(n => n % 2 == 0);
var squared = evens.Select(n => n * n);
Console.WriteLine("Done");

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
