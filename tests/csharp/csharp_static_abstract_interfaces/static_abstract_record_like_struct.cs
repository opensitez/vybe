// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_record_like_struct
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IRecord<T> where T:IRecord<T>{static abstract T Create(int id,string name);}
struct User:IRecord<User>{public int Id; public string Name; public static User Create(int id,string name)=>new User{Id=id,Name=name};}
__Check((User.Create(1,"Ann").Name).ToString(), "Ann");
