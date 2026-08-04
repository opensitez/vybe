// vybe-test: csharp/csharp_linq_groupby_join/join_links_two_sequences_on_matching_keys
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

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

var ids  = new[] { 1, 2, 3 };
var names = new[] { (Id:1, Name:"one"), (Id:2, Name:"two") };
var joined = ids.Join(names, id => id, n => n.Id, (id, n) => n.Name);
foreach (var s in joined) __P((s).ToString());
__Check("one\ntwo");
