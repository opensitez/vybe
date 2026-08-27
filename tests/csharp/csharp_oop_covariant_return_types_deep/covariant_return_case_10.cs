// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_10

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

var repo = new DerivedRepo_10();
__P((repo.Get() is DerivedEntity_10).ToString());
__Check("True");

class BaseEntity_10 { }
class DerivedEntity_10 : BaseEntity_10 { }
abstract class BaseRepo_10 {
    public abstract BaseEntity_10 Get();
}
class DerivedRepo_10 : BaseRepo_10 {
    public override DerivedEntity_10 Get() => new DerivedEntity_10();
}
