// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_names_returns_all_identifiers
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

enum Coin{Penny,Nickel,Dime} foreach(var name in System.Enum.GetNames(typeof(Coin))) Console.WriteLine(name);
