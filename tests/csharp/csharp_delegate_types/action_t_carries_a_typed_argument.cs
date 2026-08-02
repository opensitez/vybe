// vybe-test: csharp/csharp_delegate_types/action_t_carries_a_typed_argument
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action<int> print = n => __Check((n * 2).ToString(), "10");
print(5);
