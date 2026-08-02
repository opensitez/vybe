// vybe-test: csharp/csharp_inheritance_virtual_dispatch/inheritance_virtual_dispatch_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_inheritance_virtual_dispatch.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// inheritance_virtual_dispatch
int? maybe = null; int fallback = maybe ?? 71; __Check((fallback == 71).ToString(), "True");
