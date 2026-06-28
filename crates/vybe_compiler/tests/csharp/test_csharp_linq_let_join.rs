//! LINQ query syntax: `let`, `join`, `group by`, `into`, subqueries.
use super::helpers::run_csharp;

#[test]
fn let_clause_introduces_named_subexpression() {
    assert_eq!(
        run_csharp(
            r#"var result =
    from s in new[]{"hello","hi","world"}
    let len=s.Length
    where len>3
    select s;
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["hello", "world"]
    );
}

#[test]
fn join_correlates_two_sequences_on_matching_key() {
    assert_eq!(
        run_csharp(
            r#"var ids=new[]{1,2,3};
var labels=new[]{(Id:1,Text:"one"),(Id:2,Text:"two")};
var q=from id in ids
      join l in labels on id equals l.Id
      select l.Text;
foreach(var x in q) Console.WriteLine(x);"#
        ),
        &["one", "two"]
    );
}

#[test]
fn group_by_query_syntax_groups_by_first_char() {
    assert_eq!(
        run_csharp(
            r#"var words=new[]{"apple","ant","banana"};
var groups=from w in words group w by w[0];
int count=0;
foreach(var g in groups) count++;
Console.WriteLine(count);"#
        ),
        &["2"]
    );
}

#[test]
fn into_continues_query_after_group() {
    assert_eq!(
        run_csharp(
            r#"var nums=new[]{1,2,3,4,5,6};
var q=from n in nums
      group n by n%2 into g
      select g.Key;
var keys=q.OrderBy(x=>x).ToList();
Console.WriteLine(keys[0]); Console.WriteLine(keys[1]);"#
        ),
        &["0", "1"]
    );
}

#[test]
fn multiple_from_clauses_produce_cartesian_product() {
    assert_eq!(
        run_csharp(
            r#"var result=from a in new[]{1,2} from b in new[]{10,20} select a*b;
int sum=0; foreach(var x in result) sum+=x;
Console.WriteLine(sum);"#
        ),
        &["60"]
    );
}

#[test]
fn order_by_in_query_syntax_sorts_ascending() {
    assert_eq!(
        run_csharp(
            r#"var q=from n in new[]{3,1,2} orderby n select n;
foreach(var x in q) Console.WriteLine(x);"#
        ),
        &["1", "2", "3"]
    );
}
