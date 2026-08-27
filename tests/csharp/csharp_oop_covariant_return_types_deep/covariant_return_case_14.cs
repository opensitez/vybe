// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_14

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

var repo = new DerivedRepo_14();
__P((repo.Get() is DerivedEntity_14).ToString());
__Check("True");

class BaseEntity_14 { }
class DerivedEntity_14 : BaseEntity_14 { }
abstract class BaseRepo_14 {
    public abstract BaseEntity_14 Get();
}
class DerivedRepo_14 : BaseRepo_14 {
    public override DerivedEntity_14 Get() => new DerivedEntity_14();
}
