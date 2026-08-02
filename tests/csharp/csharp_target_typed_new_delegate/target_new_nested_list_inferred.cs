// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_nested_list_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<System.Collections.Generic.List<int>> grid = new();
System.Collections.Generic.List<int> row = new() { 1, 2 };
grid.Add(row);
__Check((grid[0][1]).ToString(), "2");
