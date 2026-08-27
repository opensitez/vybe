// vybe-test: csharp/csharp_linq_order_and_order_descending/linq_order_case_15

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

int[] nums = new int[] { 5, 1, 4, 2 };
var sorted = nums.Order().ToList();
__P(sorted[0].ToString());
__P(sorted[3].ToString());
__Check("1\n5");
