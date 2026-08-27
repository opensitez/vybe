// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_15

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

var repo = new DerivedRepo_15();
__P((repo.Get() is DerivedEntity_15).ToString());
__Check("True");

class BaseEntity_15 { }
class DerivedEntity_15 : BaseEntity_15 { }
abstract class BaseRepo_15 {
    public abstract BaseEntity_15 Get();
}
class DerivedRepo_15 : BaseRepo_15 {
    public override DerivedEntity_15 Get() => new DerivedEntity_15();
}
