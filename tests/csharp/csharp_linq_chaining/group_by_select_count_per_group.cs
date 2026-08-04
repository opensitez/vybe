// vybe-test: csharp/csharp_linq_chaining/group_by_select_count_per_group
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

var words=new[]{"cat","car","bar","bat","can"};
var groups=words.GroupBy(w=>w[0])
    .Select(g=>(g.Key,g.Count()))
    .OrderBy(t=>t.Key);
foreach(var(k,c) in groups) __P(($"{k}:{c}").ToString());
__Check("b:2\nc:3");
