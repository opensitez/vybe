// vybe-test: csharp/csharp_exceptions_hierarchy/custom_exception_stores_custom_message
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

class AppEx:System.Exception{public AppEx(string m):base(m){}}
string r="";
try{throw new AppEx("fail");}
catch(AppEx ex){r=ex.Message;}
__P((r).ToString());
__Check("fail");
