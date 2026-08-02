// vybe-test: csharp/csharp_numeric_precision/big_integer_arithmetic_exact_for_large_factorial
// origin: languages/csharp/tests/csharp/test_csharp_numeric_precision.rs

System.Numerics.BigInteger f=1;
for(int i=1;i<=20;i++) f*=i;
Console.WriteLine(f.ToString());
