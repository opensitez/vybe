// vybe-test: csharp/csharp_generics_constraints/generic_method_can_swap_two_values_by_tuple_assignment
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(T, T) Swap<T>(T left, T right) { (left, right) = (right, left); return (left, right); } var result = Swap(1, 9); __Check((result.Item1).ToString(), "9"); __Check((result.Item2).ToString(), "1");
