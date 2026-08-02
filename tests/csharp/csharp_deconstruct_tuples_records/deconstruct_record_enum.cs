// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_record_enum
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode{Off,On} record State(Mode M); var (m)=new State(Mode.On); __Check((m).ToString(), "On");
