// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_18

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

var repo = new DerivedRepo_18();
__P((repo.Get() is DerivedEntity_18).ToString());
__Check("True");

class BaseEntity_18 { }
class DerivedEntity_18 : BaseEntity_18 { }
abstract class BaseRepo_18 {
    public abstract BaseEntity_18 Get();
}
class DerivedRepo_18 : BaseRepo_18 {
    public override DerivedEntity_18 Get() => new DerivedEntity_18();
}
