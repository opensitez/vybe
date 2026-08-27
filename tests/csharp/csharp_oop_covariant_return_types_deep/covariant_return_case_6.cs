// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_6

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

var repo = new DerivedRepo_6();
__P((repo.Get() is DerivedEntity_6).ToString());
__Check("True");

class BaseEntity_6 { }
class DerivedEntity_6 : BaseEntity_6 { }
abstract class BaseRepo_6 {
    public abstract BaseEntity_6 Get();
}
class DerivedRepo_6 : BaseRepo_6 {
    public override DerivedEntity_6 Get() => new DerivedEntity_6();
}
