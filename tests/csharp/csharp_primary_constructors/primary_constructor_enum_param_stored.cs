// vybe-test: csharp/csharp_primary_constructors/primary_constructor_enum_param_stored
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Level { Low, High }
class Job(Level tier) { public Level Tier => tier; }
__Check((new Job(Level.High).Tier).ToString(), "High");
