// vybe-test: csharp/csharp_generic_inference_calls/covariant_array_assignment_allows_derived_elements_in_object_array
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Fruit { public string Name; }
class Apple : Fruit { }
Fruit[] basket = new Apple[2];
basket[0] = new Apple { Name = "fuji" };
__Check((basket[0].Name).ToString(), "fuji");
