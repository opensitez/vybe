// vybe-test: csharp/csharp_comparison_sorting/icomparable_implementation_drives_default_sort_order
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using System.Collections.Generic; class Rank : System.IComparable<Rank> { public int Value; public int CompareTo(Rank other) { return Value.CompareTo(other.Value); } } var list = new List<Rank> { new Rank { Value = 3 }, new Rank { Value = 1 } }; list.Sort(); foreach (var item in list) Console.WriteLine(item.Value);
