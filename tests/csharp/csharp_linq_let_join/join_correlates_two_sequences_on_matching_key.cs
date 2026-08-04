// vybe-test: csharp/csharp_linq_let_join/join_correlates_two_sequences_on_matching_key
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

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

var ids=new[]{1,2,3};
var labels=new[]{(Id:1,Text:"one"),(Id:2,Text:"two")};
var q=from id in ids
      join l in labels on id equals l.Id
      select l.Text;
foreach(var x in q) __P((x).ToString());
__Check("one\ntwo");
