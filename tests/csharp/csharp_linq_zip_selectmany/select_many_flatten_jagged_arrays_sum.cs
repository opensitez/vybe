// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_flatten_jagged_arrays_sum
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

var flat=new[]{new[]{1,2},new[]{3,4,5}}.SelectMany(x=>x);
__P((flat.Sum()).ToString());
__Check("15");
