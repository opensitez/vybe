// vybe-test: csharp/csharp_linq_numeric/average_of_integer_sequence_returns_double
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double avg=new[]{1,2,3,4,5}.Average();
__Check((avg).ToString(), "3");
