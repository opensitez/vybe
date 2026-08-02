// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_tuple_from_array
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var arr=new[]{(1,2),(3,4)}; var (x,y)=arr[1]; __Check((x+y).ToString(), "7");
