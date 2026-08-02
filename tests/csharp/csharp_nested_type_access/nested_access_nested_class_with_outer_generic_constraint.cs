// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_with_outer_generic_constraint
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Repo<T>{public class Row{public T Data;} public T Read(Row r)=>r.Data;} __Check((new Repo<int>().Read(new Repo<int>.Row{Data=77})).ToString(), "77");
