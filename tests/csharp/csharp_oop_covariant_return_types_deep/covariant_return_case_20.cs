// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_20

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

var repo = new DerivedRepo_20();
__P((repo.Get() is DerivedEntity_20).ToString());
__Check("True");

class BaseEntity_20 { }
class DerivedEntity_20 : BaseEntity_20 { }
abstract class BaseRepo_20 {
    public abstract BaseEntity_20 Get();
}
class DerivedRepo_20 : BaseRepo_20 {
    public override DerivedEntity_20 Get() => new DerivedEntity_20();
}
