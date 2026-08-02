// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
int seed = 71; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
