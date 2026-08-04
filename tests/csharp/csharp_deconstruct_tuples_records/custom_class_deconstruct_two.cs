// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_class_deconstruct_two
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Size{public int W,H; public void Deconstruct(out int w,out int h){w=W;h=H;}} var (w,h)=new Size{W=3,H=4}; __P((w+h).ToString());
__Check("7");
