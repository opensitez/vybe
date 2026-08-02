// vybe-test: csharp/csharp_random_random/seeded_random_produces_deterministic_sequence
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r1=new System.Random(99); var r2=new System.Random(99);
__Check((r1.Next()==r2.Next()).ToString(), "True");
