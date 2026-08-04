// vybe-test: csharp/csharp_generic_inference_calls/generic_method_with_multiple_type_params_infers_both
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

(K, V) MakePair<K, V>(K key, V value) { return (key, value); }
var pair = MakePair("id", 7);
__P((pair.Item1).ToString());
__P((pair.Item2).ToString());
__Check("id\n7");
