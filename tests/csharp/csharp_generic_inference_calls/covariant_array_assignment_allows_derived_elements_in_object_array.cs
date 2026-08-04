// vybe-test: csharp/csharp_generic_inference_calls/covariant_array_assignment_allows_derived_elements_in_object_array
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

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

class Fruit { public string Name; }
class Apple : Fruit { }
Fruit[] basket = new Apple[2];
basket[0] = new Apple { Name = "fuji" };
__P((basket[0].Name).ToString());
__Check("fuji");
