// vybe-test: csharp/csharp_anonymous_types/anonymous_type_from_linq_select_projection
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=new[]{(Id:1,Name:"a"),(Id:2,Name:"b")};
var result=data.Select(d=>new{d.Id,Upper=d.Name.ToUpper()}).ToList();
__Check((result[1].Upper).ToString(), "B");
