// vybe-test: csharp/csharp_constructor_chains/struct_constructor_can_delegate_to_field_assignment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Pair { public int Left; public int Right; public Pair(int left, int right) { Left = left; Right = right; } } var pair = new Pair(2, 8); __Check((pair.Left + pair.Right).ToString(), "10");
