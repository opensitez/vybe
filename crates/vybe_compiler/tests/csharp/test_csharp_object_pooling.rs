//! `Microsoft.Extensions.ObjectPool` pattern simulation with custom pool.
use super::helpers::run_csharp;

#[test]
fn simple_object_pool_reuses_instances() {
    assert_eq!(
        run_csharp(r#"class Pool<T> where T:new(){
    System.Collections.Generic.Queue<T> _q=new();
    public T Get()=>_q.Count>0?_q.Dequeue():new T();
    public void Return(T t)=>_q.Enqueue(t);
}
class Widget{public int Id;}
var pool=new Pool<Widget>();
var w1=pool.Get(); w1.Id=1;
pool.Return(w1);
var w2=pool.Get();
Console.WriteLine(w2.Id);"#),
        &["1"]
    );
}

#[test]
fn pool_creates_new_when_empty() {
    assert_eq!(
        run_csharp(r#"class Pool<T> where T:new(){
    System.Collections.Generic.Queue<T> _q=new();
    public T Get()=>_q.Count>0?_q.Dequeue():new T();
    public void Return(T t)=>_q.Enqueue(t);
}
class Counter{public int V=0;}
var pool=new Pool<Counter>();
var c=pool.Get();
Console.WriteLine(c.V);"#),
        &["0"]
    );
}
