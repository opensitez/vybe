// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_record_like_struct
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

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

interface IRecord<T> where T:IRecord<T>{static abstract T Create(int id,string name);}
struct User:IRecord<User>{public int Id; public string Name; public static User Create(int id,string name)=>new User{Id=id,Name=name};}
__P((User.Create(1,"Ann").Name).ToString());
__Check("Ann");
