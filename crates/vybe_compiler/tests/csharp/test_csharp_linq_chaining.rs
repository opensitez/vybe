//! LINQ method chaining: Where+Select+OrderBy+Take, complex projections.
use super::helpers::run_csharp;

#[test]
fn where_select_order_take_pipeline() {
    assert_eq!(
        run_csharp(r#"var result=new[]{5,3,8,1,9,2,7,4,6}
    .Where(n=>n>3)
    .Select(n=>n*n)
    .OrderBy(n=>n)
    .Take(3);
foreach(var n in result) Console.WriteLine(n);"#),
        &["16", "25", "36"]
    );
}

#[test]
fn group_by_select_count_per_group() {
    assert_eq!(
        run_csharp(r#"var words=new[]{"cat","car","bar","bat","can"};
var groups=words.GroupBy(w=>w[0])
    .Select(g=>(g.Key,g.Count()))
    .OrderBy(t=>t.Key);
foreach(var(k,c) in groups) Console.WriteLine($"{k}:{c}");"#),
        &["b:2", "c:3"]
    );
}

#[test]
fn select_many_with_index_flattens_and_annotates() {
    assert_eq!(
        run_csharp(r#"var groups=new[]{new[]{1,2},new[]{3,4}};
var result=groups.SelectMany((g,i)=>g.Select(x=>i*10+x));
Console.WriteLine(string.Join(",",result));"#),
        &["1,2,13,14"]
    );
}

#[test]
fn chained_order_by_then_by_sorts_on_two_keys() {
    assert_eq!(
        run_csharp(r#"var data=new[]{(A:"b",B:2),(A:"a",B:3),(A:"a",B:1)};
var result=data.OrderBy(x=>x.A).ThenBy(x=>x.B);
foreach(var(a,b) in result) Console.WriteLine($"{a}{b}");"#),
        &["a1", "a3", "b2"]
    );
}

#[test]
fn zip_three_sequences_pairwise_sum() {
    assert_eq!(
        run_csharp(r#"var a=new[]{1,2,3}; var b=new[]{10,20,30};
var result=a.Zip(b).Select(t=>t.First+t.Second);
Console.WriteLine(string.Join(",",result));"#),
        &["11,22,33"]
    );
}
