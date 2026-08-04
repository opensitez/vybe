// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_record_style_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Orders{public class Line{public int Qty; public int Total()=>Qty*2;} public Line Make(int q)=>new Line{Qty=q};} __P((new Orders().Make(4).Total()).ToString());
__Check("8");
