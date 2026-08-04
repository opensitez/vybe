// vybe-test: csharp/csharp_linq_zip_selectmany/zip_string_pairs_second_concatenation
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

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

var z=new[]{"a","b"}.Zip(new[]{"1","2"},(x,y)=>x+y);
__P((z.Last()).ToString());
__Check("b2");
