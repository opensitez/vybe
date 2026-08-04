// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_implements_outer_interface
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Host{public interface IRun{int Go();} public class Worker:IRun{public int Go()=>4;}} __P((new Host.Worker().Go()).ToString());
__Check("4");
