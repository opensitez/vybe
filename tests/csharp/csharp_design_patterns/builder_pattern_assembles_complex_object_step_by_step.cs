// vybe-test: csharp/csharp_design_patterns/builder_pattern_assembles_complex_object_step_by_step
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

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

class Query{public string Table="";public string Filter="";}
class QueryBuilder{
    Query q=new Query();
    public QueryBuilder From(string t){q.Table=t;return this;}
    public QueryBuilder Where(string f){q.Filter=f;return this;}
    public Query Build()=>q;
}
var q=new QueryBuilder().From("users").Where("age>18").Build();
__P((q.Table).ToString()); __P((q.Filter).ToString());
__Check("users\nage>18");
