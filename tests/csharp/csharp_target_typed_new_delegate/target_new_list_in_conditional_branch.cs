// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_in_conditional_branch
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> pick(bool flag) {
    System.Collections.Generic.List<int> a = new() { 1 };
    System.Collections.Generic.List<int> b = new() { 2 };
    return flag ? a : b;
}
__Check((pick(false)[0]).ToString(), "2");
