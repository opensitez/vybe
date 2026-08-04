// vybe-test: csharp/csharp_object_pooling/simple_object_pool_reuses_instances
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
class Widget{public int Id;}
var pool=new Pool<Widget>();
var w1=pool.Get(); w1.Id=1;
pool.Return(w1);
var w2=pool.Get();
__P((w2.Id).ToString());
__Check("1");
