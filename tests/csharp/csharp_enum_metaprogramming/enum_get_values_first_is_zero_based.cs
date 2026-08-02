// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_first_is_zero_based
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

enum Rank{First,Second} foreach(var v in System.Enum.GetValues(typeof(Rank))) Console.WriteLine((int)v);
