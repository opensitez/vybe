// vybe-test: csharp/csharp_generics_where/where_new_constraint_allows_parameterless_construction
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Build<T>() where T:new()=>new T();
class Box{public int V=7;}
__Check((Build<Box>().V).ToString(), "7");
