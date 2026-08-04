// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_with_multiple_readonly_fields
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

readonly ref struct Rect{public readonly int W; public readonly int H; public Rect(int w,int h){W=w; H=h;} public int Area()=>W*H;} var r=new Rect(3,4); __P((r.Area()).ToString());
__Check("12");
