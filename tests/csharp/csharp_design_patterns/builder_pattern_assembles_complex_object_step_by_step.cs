// vybe-test: csharp/csharp_design_patterns/builder_pattern_assembles_complex_object_step_by_step
// origin: languages/csharp/tests/csharp/test_csharp_design_patterns.rs

using static __Harness;

var q=new QueryBuilder().From("users").Where("age>18").Build();
__P((q.Table).ToString());
__P((q.Filter).ToString());
__Check("users\nage>18");

class Query{public string Table="";public string Filter="";}

class QueryBuilder{
    Query q=new Query();
    public QueryBuilder From(string t){q.Table=t;return this;}
    public QueryBuilder Where(string f){q.Filter=f;return this;}
    public Query Build()=>q;
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
