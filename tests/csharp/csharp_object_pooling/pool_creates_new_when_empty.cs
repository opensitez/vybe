// vybe-test: csharp/csharp_object_pooling/pool_creates_new_when_empty
// origin: languages/csharp/tests/csharp/test_csharp_object_pooling.rs

using static __Harness;

var pool=new Pool<Counter>();
var c=pool.Get();
__P((c.V).ToString());
__Check("0");

class Pool<T> where T:new(){
    System.Collections.Generic.Queue<T> _q=new();
    public T Get()=>_q.Count>0?_q.Dequeue():new T();
    public void Return(T t)=>_q.Enqueue(t);
}

class Counter{public int V=0;}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
