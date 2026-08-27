// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_4

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

var repo = new DerivedRepo_4();
__P((repo.Get() is DerivedEntity_4).ToString());
__Check("True");

class BaseEntity_4 { }
class DerivedEntity_4 : BaseEntity_4 { }
abstract class BaseRepo_4 {
    public abstract BaseEntity_4 Get();
}
class DerivedRepo_4 : BaseRepo_4 {
    public override DerivedEntity_4 Get() => new DerivedEntity_4();
}
