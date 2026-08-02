// vybe-test: csharp/csharp_datetime_advanced/datetime_add_days_correctly_crosses_month_boundary
// origin: languages/csharp/tests/csharp/test_csharp_datetime_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var d=new System.DateTime(2024,1,30).AddDays(3);
__Check((d.Month).ToString(), "2"); __Check((d.Day).ToString(), "2");
