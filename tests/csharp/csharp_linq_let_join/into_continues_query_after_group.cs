// vybe-test: csharp/csharp_linq_let_join/into_continues_query_after_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var nums=new[]{1,2,3,4,5,6};
var q=from n in nums
      group n by n%2 into g
      select g.Key;
var keys=q.OrderBy(x=>x).ToList();
__Check((keys[0]).ToString(), "0"); __Check((keys[1]).ToString(), "1");
