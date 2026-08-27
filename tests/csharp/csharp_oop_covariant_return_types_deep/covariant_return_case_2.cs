// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_2

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var repo = new DerivedRepo_2();
__P((repo.Get() is DerivedEntity_2).ToString());
__Check("True");

class BaseEntity_2 { }
class DerivedEntity_2 : BaseEntity_2 { }
abstract class BaseRepo_2 {
    public abstract BaseEntity_2 Get();
}
class DerivedRepo_2 : BaseRepo_2 {
    public override DerivedEntity_2 Get() => new DerivedEntity_2();
}
