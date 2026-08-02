// vybe-test: csharp/csharp_params_optional_named/params_can_be_called_with_explicit_array
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
Console.WriteLine(Sum(new int[]{10,20}));
