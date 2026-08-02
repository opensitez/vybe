// vybe-test: csharp/csharp_ref_return_semantics/ref_return_from_local_function_updates_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int total = 5;
ref int Bump() => ref total;
ref int view = ref Bump();
view += 2;
__Check((total).ToString(), "7");
