// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_1

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

var repo = new DerivedRepo_1();
__P((repo.Get() is DerivedEntity_1).ToString());
__Check("True");

class BaseEntity_1 { }
class DerivedEntity_1 : BaseEntity_1 { }
abstract class BaseRepo_1 {
    public abstract BaseEntity_1 Get();
}
class DerivedRepo_1 : BaseRepo_1 {
    public override DerivedEntity_1 Get() => new DerivedEntity_1();
}
