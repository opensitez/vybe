// vybe-test: csharp/csharp_linq_projections/to_list_materializes_query_to_mutable_list
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

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

var list = new[]{1,2,3}.Select(x => x*2).ToList();
__P((list.GetType().Name).ToString());
__Check("List`1");
