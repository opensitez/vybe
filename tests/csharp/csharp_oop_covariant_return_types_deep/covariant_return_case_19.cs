// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_19

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

var repo = new DerivedRepo_19();
__P((repo.Get() is DerivedEntity_19).ToString());
__Check("True");

class BaseEntity_19 { }
class DerivedEntity_19 : BaseEntity_19 { }
abstract class BaseRepo_19 {
    public abstract BaseEntity_19 Get();
}
class DerivedRepo_19 : BaseRepo_19 {
    public override DerivedEntity_19 Get() => new DerivedEntity_19();
}
