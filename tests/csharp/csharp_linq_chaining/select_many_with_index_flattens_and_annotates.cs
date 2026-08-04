// vybe-test: csharp/csharp_linq_chaining/select_many_with_index_flattens_and_annotates
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

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

var groups=new[]{new[]{1,2},new[]{3,4}};
var result=groups.SelectMany((g,i)=>g.Select(x=>i*10+x));
__P((string.Join(",",result)).ToString());
__Check("1,2,13,14");
