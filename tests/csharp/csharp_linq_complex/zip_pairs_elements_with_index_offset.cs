// vybe-test: csharp/csharp_linq_complex/zip_pairs_elements_with_index_offset
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

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

var a=new[]{1,2,3}; var b=new[]{4,5,6};
var r=a.Zip(b,(x,y)=>x*y);
__P((r.Sum()).ToString());
__Check("32");
