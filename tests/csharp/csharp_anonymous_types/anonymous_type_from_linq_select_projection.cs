// vybe-test: csharp/csharp_anonymous_types/anonymous_type_from_linq_select_projection
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_types.rs

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

var data=new[]{(Id:1,Name:"a"),(Id:2,Name:"b")};
var result=data.Select(d=>new{d.Id,Upper=d.Name.ToUpper()}).ToList();
__P((result[1].Upper).ToString());
__Check("B");
