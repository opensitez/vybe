// vybe-test: csharp/csharp_generic_methods/generic_method_filters_sequence_by_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

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

System.Collections.Generic.IEnumerable<T> FilterType<T>(object[] items){
    foreach(var i in items) if(i is T t) yield return t;
}
var items=new object[]{1,"a",2,"b",3};
int count=0;
foreach(var s in FilterType<string>(items)) count++;
__P((count).ToString());
__Check("2");
