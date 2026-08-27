// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_17

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

var repo = new DerivedRepo_17();
__P((repo.Get() is DerivedEntity_17).ToString());
__Check("True");

class BaseEntity_17 { }
class DerivedEntity_17 : BaseEntity_17 { }
abstract class BaseRepo_17 {
    public abstract BaseEntity_17 Get();
}
class DerivedRepo_17 : BaseRepo_17 {
    public override DerivedEntity_17 Get() => new DerivedEntity_17();
}
