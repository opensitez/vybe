// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_build_comma_list_value
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var text=new[]{1,2,3}.Aggregate("",(acc,x)=>acc==""?x.ToString():acc+","+x);
__P((text).ToString());
__Check("1,2,3");
