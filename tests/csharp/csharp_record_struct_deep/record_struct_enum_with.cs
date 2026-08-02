// vybe-test: csharp/csharp_record_struct_deep/record_struct_enum_with
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Level{Low,High} record struct Job(Level Tier); var j=new Job(Level.Low); var k=j with{Tier=Level.High}; __Check((k.Tier).ToString(), "High");
