// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_implements_outer_interface
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Host{public interface IRun{int Go();} public class Worker:IRun{public int Go()=>4;}} __Check((new Host.Worker().Go()).ToString(), "4");
