// vybe-test: csharp/csharp_linq_take_range_and_skiplast/linq_take_range_case_4

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

int[] nums = new int[] { 10, 20, 30, 40, 50 };
var slice = nums.Take(1..4).ToList();
__P(slice.Count.ToString());
__P(slice[0].ToString());
__Check("3\n20");
