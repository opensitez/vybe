// vybe-test: csharp/csharp_design_patterns/builder_pattern_assembles_complex_object_step_by_step
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((q.Table).ToString(), "users"); __Check((q.Filter).ToString(), "age>18");
