// vybe-test: csharp/csharp_delegate_types/action_delegate_calls_void_method_with_no_args
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Action greet = () => __Check(("hi").ToString(), "hi");
greet();
