// vybe-test: csharp/csharp_ref_return_semantics/ref_return_chains_to_second_ref_local_without_copying_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] values = { 10, 20 };
ref int First() => ref values[0];
ref int alias = ref First();
alias = 99;
__Check((values[0]).ToString(), "99");
