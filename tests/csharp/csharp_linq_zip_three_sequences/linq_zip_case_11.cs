// vybe-test: csharp/csharp_linq_zip_three_sequences/linq_zip_case_11

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

int[] s1 = new int[] { 11 };
string[] s2 = new string[] { "A" };
bool[] s3 = new bool[] { true };
var zipped = System.Linq.Enumerable.Zip(s1, s2, s3).ToList();
__P(zipped.Count.ToString());
__P(zipped[0].First.ToString());
__P(zipped[0].Second);
__Check("1\n11\nA");
