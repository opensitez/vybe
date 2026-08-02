// vybe-test: csharp/csharp_delegate_types/removing_handler_from_multicast_leaves_remaining
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int count = 0;
System.Action a = () => count++;
System.Action b = () => count++;
System.Action multi = a;
multi += b;
multi -= a;
multi();
__Check((count).ToString(), "1");
