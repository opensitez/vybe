// vybe-test: csharp/csharp_generic_inference_calls/generic_method_with_multiple_type_params_infers_both
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(K, V) MakePair<K, V>(K key, V value) { return (key, value); }
var pair = MakePair("id", 7);
__Check((pair.Item1).ToString(), "id");
__Check((pair.Item2).ToString(), "7");
