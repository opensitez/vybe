// vybe-test: csharp/csharp_exceptions_hierarchy/catch_base_exception_catches_derived_type
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_hierarchy.rs

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

string r="";
try{int[] a=new int[3]; var _=a[10];}
catch(System.Exception ex){r=ex.GetType().Name;}
__P((r).ToString());
__Check("IndexOutOfRangeException");
