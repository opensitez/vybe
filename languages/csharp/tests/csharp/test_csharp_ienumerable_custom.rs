//! Custom `IEnumerable<T>` and `IEnumerator<T>` implementations.
use super::helpers::run_csharp;

#[test]
fn custom_enumerable_iterated_by_foreach() {
    assert_eq!(
        run_csharp(
            r#"class UpTo:System.Collections.Generic.IEnumerable<int>{
    int _max;
    public UpTo(int max){_max=max;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=1;i<=_max;i++) yield return i;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}
int sum=0; foreach(var n in new UpTo(5)) sum+=n;
Console.WriteLine(sum);"#
        ),
        &["15"]
    );
}

#[test]
fn manual_enumerator_move_next_and_current() {
    assert_eq!(
        run_csharp(
            r#"var list=new System.Collections.Generic.List<int>{10,20,30};
using var e=list.GetEnumerator();
e.MoveNext();
Console.WriteLine(e.Current);"#
        ),
        &["10"]
    );
}

#[test]
fn linq_works_on_custom_ienumerable() {
    assert_eq!(
        run_csharp(
            r#"class Odds:System.Collections.Generic.IEnumerable<int>{
    int _count;
    public Odds(int count){_count=count;}
    public System.Collections.Generic.IEnumerator<int> GetEnumerator(){
        for(int i=0;i<_count;i++) yield return 2*i+1;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()=>GetEnumerator();
}
Console.WriteLine(new Odds(4).Sum());"#
        ),
        &["16"]
    );
}

#[test]
fn reset_on_list_enumerator_restarts_sequence() {
    assert_eq!(
        run_csharp(
            r#"var list=new System.Collections.Generic.List<int>{1,2,3};
int count=0;
foreach(var _ in list) count++;
foreach(var _ in list) count++;
Console.WriteLine(count);"#
        ),
        &["6"]
    );
}
