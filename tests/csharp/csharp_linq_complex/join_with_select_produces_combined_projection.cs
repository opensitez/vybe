// vybe-test: csharp/csharp_linq_complex/join_with_select_produces_combined_projection
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

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

var users=new[]{(Id:1,Name:"Alice"),(Id:2,Name:"Bob")};
var orders=new[]{(UserId:1,Item:"book"),(UserId:2,Item:"pen"),(UserId:1,Item:"cup")};
var q=orders.Join(users,o=>o.UserId,u=>u.Id,(o,u)=>$"{u.Name}:{o.Item}")
    .OrderBy(s=>s);
foreach(var s in q) __P((s).ToString());
__Check("Alice:book\nAlice:cup\nBob:pen");
