// vybe-test: csharp/csharp_enum_metaprogramming/enum_get_values_cast_to_int_array
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

enum Score{A=1,B=3,C=5} int sum=0; foreach(var v in System.Enum.GetValues(typeof(Score))) sum+=(int)v; Console.WriteLine(sum);
