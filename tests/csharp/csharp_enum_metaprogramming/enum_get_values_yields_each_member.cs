// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_yields_each_member
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

enum Pair{A,B,C} foreach(var v in System.Enum.GetValues(typeof(Pair))) Console.WriteLine(v);
