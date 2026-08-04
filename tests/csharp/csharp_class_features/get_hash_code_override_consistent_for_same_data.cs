// vybe-test: csharp/csharp_class_features/get_hash_code_override_consistent_for_same_data
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Key{int V;public Key(int v){V=v;}public override int GetHashCode()=>V.GetHashCode();}
__P((new Key(7).GetHashCode()==new Key(7).GetHashCode()).ToString());
__Check("True");
