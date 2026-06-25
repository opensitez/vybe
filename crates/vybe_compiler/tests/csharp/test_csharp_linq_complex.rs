//! Complex LINQ: chained groupby+orderby, join+select, SelectMany flat maps.
use super::helpers::run_csharp;

#[test]
fn group_by_then_order_group_keys_ascending() {
    assert_eq!(
        run_csharp(r#"var items=new[]{(Cat:"b",Val:2),(Cat:"a",Val:1),(Cat:"b",Val:4),(Cat:"a",Val:3)};
var groups=items.GroupBy(i=>i.Cat).OrderBy(g=>g.Key)
    .Select(g=>(g.Key,g.Sum(i=>i.Val)));
foreach(var(k,s) in groups) Console.WriteLine($"{k}:{s}");"#),
        &["a:4", "b:6"]
    );
}

#[test]
fn join_with_select_produces_combined_projection() {
    assert_eq!(
        run_csharp(r#"var users=new[]{(Id:1,Name:"Alice"),(Id:2,Name:"Bob")};
var orders=new[]{(UserId:1,Item:"book"),(UserId:2,Item:"pen"),(UserId:1,Item:"cup")};
var q=orders.Join(users,o=>o.UserId,u=>u.Id,(o,u)=>$"{u.Name}:{o.Item}")
    .OrderBy(s=>s);
foreach(var s in q) Console.WriteLine(s);"#),
        &["Alice:book", "Alice:cup", "Bob:pen"]
    );
}

#[test]
fn select_many_flattens_nested_lists() {
    assert_eq!(
        run_csharp(r#"var data=new[]{
    new[]{1,2,3},
    new[]{4,5},
    new[]{6}
};
int sum=data.SelectMany(x=>x).Sum();
Console.WriteLine(sum);"#),
        &["21"]
    );
}

#[test]
fn zip_pairs_elements_with_index_offset() {
    assert_eq!(
        run_csharp(r#"var a=new[]{1,2,3}; var b=new[]{4,5,6};
var r=a.Zip(b,(x,y)=>x*y);
Console.WriteLine(r.Sum());"#),
        &["32"]
    );
}

#[test]
fn aggregate_with_seed_computes_running_product() {
    assert_eq!(
        run_csharp(r#"var result=new[]{1,2,3,4,5}.Aggregate(1L,(acc,n)=>acc*n);
Console.WriteLine(result);"#),
        &["120"]
    );
}

#[test]
fn to_lookup_groups_items_by_key_returning_ilookup() {
    assert_eq!(
        run_csharp(r#"var data=new[]{(K:"a",V:1),(K:"a",V:2),(K:"b",V:3)};
var lu=data.ToLookup(x=>x.K,x=>x.V);
Console.WriteLine(lu["a"].Sum());
Console.WriteLine(lu["b"].Sum());"#),
        &["3", "3"]
    );
}
