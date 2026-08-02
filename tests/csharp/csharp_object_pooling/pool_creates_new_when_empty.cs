// vybe-test: csharp/csharp_object_pooling/pool_creates_new_when_empty
// origin: languages/csharp/tests/csharp/test_csharp_object_pooling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pool<T> where T:new(){
    System.Collections.Generic.Queue<T> _q=new();
    public T Get()=>_q.Count>0?_q.Dequeue():new T();
    public void Return(T t)=>_q.Enqueue(t);
}
class Counter{public int V=0;}
var pool=new Pool<Counter>();
var c=pool.Get();
__Check((c.V).ToString(), "0");
