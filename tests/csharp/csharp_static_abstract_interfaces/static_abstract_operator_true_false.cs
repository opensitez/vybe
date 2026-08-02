// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_operator_true_false
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ITest<T> where T:ITest<T>{static abstract bool operator true(T v); static abstract bool operator false(T v);}
struct Flag:ITest<Flag>{public bool On; public static bool operator true(Flag v)=>v.On; public static bool operator false(Flag v)=>!v.On;}
__Check((new Flag{On=true}?1:0).ToString(), "1");
