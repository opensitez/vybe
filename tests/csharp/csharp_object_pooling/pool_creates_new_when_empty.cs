// vybe-test: csharp/csharp_object_pooling/pool_creates_new_when_empty
// origin: languages/csharp/tests/csharp/test_csharp_object_pooling.rs

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

class Pool<T> where T:new(){
    System.Collections.Generic.Queue<T> _q=new();
    public T Get()=>_q.Count>0?_q.Dequeue():new T();
    public void Return(T t)=>_q.Enqueue(t);
}
class Counter{public int V=0;}
var pool=new Pool<Counter>();
var c=pool.Get();
__P((c.V).ToString());
__Check("0");
