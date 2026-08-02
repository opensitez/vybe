// vybe-test: csharp/csharp_generic_methods/generic_method_on_non_generic_class_infers_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Utils{public static T First<T>(T[] arr)=>arr[0];}
__Check((Utils.First(new[]{10,20,30})).ToString(), "10");
__Check((Utils.First(new[]{"a","b"})).ToString(), "a");
