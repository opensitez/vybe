// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_5

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

var repo = new DerivedRepo_5();
__P((repo.Get() is DerivedEntity_5).ToString());
__Check("True");

class BaseEntity_5 { }
class DerivedEntity_5 : BaseEntity_5 { }
abstract class BaseRepo_5 {
    public abstract BaseEntity_5 Get();
}
class DerivedRepo_5 : BaseRepo_5 {
    public override DerivedEntity_5 Get() => new DerivedEntity_5();
}
