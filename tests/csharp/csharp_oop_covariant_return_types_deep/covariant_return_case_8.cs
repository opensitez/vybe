// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_8

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

var repo = new DerivedRepo_8();
__P((repo.Get() is DerivedEntity_8).ToString());
__Check("True");

class BaseEntity_8 { }
class DerivedEntity_8 : BaseEntity_8 { }
abstract class BaseRepo_8 {
    public abstract BaseEntity_8 Get();
}
class DerivedRepo_8 : BaseRepo_8 {
    public override DerivedEntity_8 Get() => new DerivedEntity_8();
}
