// vybe-test: csharp/csharp_params_optional_named/params_accepts_variable_number_of_arguments
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
Console.WriteLine(Sum(1,2,3,4,5));
