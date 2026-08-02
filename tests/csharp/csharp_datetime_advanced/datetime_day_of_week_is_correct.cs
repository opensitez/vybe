// vybe-test: csharp/csharp_datetime_advanced/datetime_day_of_week_is_correct
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=new System.DateTime(2024,1,1);
__Check((d.DayOfWeek).ToString(), "Monday");
