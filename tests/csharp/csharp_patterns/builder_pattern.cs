// vybe-test: csharp/csharp_patterns/builder_pattern
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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

class QueryBuilder {
    private string query = "SELECT *";
    public QueryBuilder From(string table) {
        query += " FROM " + table;
        return this;
    }
    public QueryBuilder Where(string condition) {
        query += " WHERE " + condition;
        return this;
    }
    public string Build() { return query; }
}
var q = new QueryBuilder().From("users").Where("age > 18").Build();
__P((q).ToString());
__Check("SELECT * FROM users WHERE age > 18");
