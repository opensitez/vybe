//! Classic OOP design patterns: Singleton, Strategy, Observer, Builder, Factory.
use super::helpers::run_csharp;

#[test]
fn singleton_returns_same_instance_on_repeated_calls() {
    assert_eq!(
        run_csharp(
            r#"class Singleton{
    static Singleton _inst;
    public int Val;
    public static Singleton Instance=>_inst??=new Singleton();
}
Singleton.Instance.Val=42;
Console.WriteLine(Singleton.Instance.Val);"#
        ),
        &["42"]
    );
}

#[test]
fn strategy_pattern_swaps_algorithm_at_runtime() {
    assert_eq!(
        run_csharp(
            r#"interface ISort{int[] Sort(int[] a);}
class Ascending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);return c;}}
class Descending:ISort{public int[] Sort(int[] a){var c=(int[])a.Clone();System.Array.Sort(c);System.Array.Reverse(c);return c;}}
ISort s=new Ascending();
Console.WriteLine(string.Join(",",s.Sort(new[]{3,1,2})));
s=new Descending();
Console.WriteLine(string.Join(",",s.Sort(new[]{3,1,2})));"#
        ),
        &["1,2,3", "3,2,1"]
    );
}

#[test]
fn observer_pattern_notifies_multiple_subscribers() {
    assert_eq!(
        run_csharp(
            r#"class Button{public event System.Action Clicked;}
var b=new Button();
int a=0,c=0;
b.Clicked+=()=>a++;
b.Clicked+=()=>c++;
b.Clicked?.Invoke();
Console.WriteLine(a); Console.WriteLine(c);"#
        ),
        &["1", "1"]
    );
}

#[test]
fn builder_pattern_assembles_complex_object_step_by_step() {
    assert_eq!(
        run_csharp(
            r#"class Query{public string Table="";public string Filter="";}
class QueryBuilder{
    Query q=new Query();
    public QueryBuilder From(string t){q.Table=t;return this;}
    public QueryBuilder Where(string f){q.Filter=f;return this;}
    public Query Build()=>q;
}
var q=new QueryBuilder().From("users").Where("age>18").Build();
Console.WriteLine(q.Table); Console.WriteLine(q.Filter);"#
        ),
        &["users", "age>18"]
    );
}

#[test]
fn factory_method_creates_correct_concrete_type() {
    assert_eq!(
        run_csharp(
            r#"abstract class Animal{public abstract string Sound();}
class Dog:Animal{public override string Sound()=>"woof";}
class Cat:Animal{public override string Sound()=>"meow";}
Animal Create(string kind)=>kind=="dog"?(Animal)new Dog():new Cat();
Console.WriteLine(Create("dog").Sound());
Console.WriteLine(Create("cat").Sound());"#
        ),
        &["woof", "meow"]
    );
}
