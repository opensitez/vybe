// vybe-test: csharp/csharp_linq_zip_selectmany/zip_result_selector_returns_tuple_like_sum
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

var z=new[]{1,2}.Zip(new[]{3,4},(a,b)=>a*10+b);
__P((z.Sum()).ToString());
__Check("47");
