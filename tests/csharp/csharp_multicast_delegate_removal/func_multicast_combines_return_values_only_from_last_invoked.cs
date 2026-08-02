// vybe-test: csharp/csharp_multicast_delegate_removal/func_multicast_combines_return_values_only_from_last_invoked
// origin: languages/csharp/tests/csharp/test_csharp_multicast_delegate_removal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System;
Func<int> first = () => { __Check(("1").ToString(), "1"); return 1; };
Func<int> second = () => { __Check(("2").ToString(), "2"); return 2; };
Func<int> chain = first;
chain += second;
__Check((chain()).ToString(), "2");
