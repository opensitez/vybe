// vybe-test: csharp/csharp_linq_complex/join_with_select_produces_combined_projection
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

var users=new[]{(Id:1,Name:"Alice"),(Id:2,Name:"Bob")};
var orders=new[]{(UserId:1,Item:"book"),(UserId:2,Item:"pen"),(UserId:1,Item:"cup")};
var q=orders.Join(users,o=>o.UserId,u=>u.Id,(o,u)=>$"{u.Name}:{o.Item}")
    .OrderBy(s=>s);
foreach(var s in q) Console.WriteLine(s);
