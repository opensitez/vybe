// vybe-test: csharp/csharp_abstract_sealed/sealed_class_cannot_be_used_as_base_detected_at_compile_time_but_runtime_ok
// origin: languages/csharp/tests/csharp/test_csharp_abstract_sealed.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

sealed class Final { public int Value = 7; }
var f = new Final();
__Check((f.Value).ToString(), "7");
