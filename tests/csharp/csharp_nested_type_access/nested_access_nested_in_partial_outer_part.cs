// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_in_partial_outer_part
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Worker{public class Helper{public int Run()=>1;}} partial class Worker{public int Go()=>new Helper().Run();} __Check((new Worker().Go()).ToString(), "1");
