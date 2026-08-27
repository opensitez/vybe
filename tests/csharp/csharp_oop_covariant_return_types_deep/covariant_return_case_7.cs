// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_7

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

var repo = new DerivedRepo_7();
__P((repo.Get() is DerivedEntity_7).ToString());
__Check("True");

class BaseEntity_7 { }
class DerivedEntity_7 : BaseEntity_7 { }
abstract class BaseRepo_7 {
    public abstract BaseEntity_7 Get();
}
class DerivedRepo_7 : BaseRepo_7 {
    public override DerivedEntity_7 Get() => new DerivedEntity_7();
}
