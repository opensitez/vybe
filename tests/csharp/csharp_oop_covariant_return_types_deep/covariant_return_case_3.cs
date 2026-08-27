// vybe-test: csharp/csharp_oop_covariant_return_types_deep/covariant_return_case_3

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

var repo = new DerivedRepo_3();
__P((repo.Get() is DerivedEntity_3).ToString());
__Check("True");

class BaseEntity_3 { }
class DerivedEntity_3 : BaseEntity_3 { }
abstract class BaseRepo_3 {
    public abstract BaseEntity_3 Get();
}
class DerivedRepo_3 : BaseRepo_3 {
    public override DerivedEntity_3 Get() => new DerivedEntity_3();
}
