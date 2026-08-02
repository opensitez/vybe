// vybe-test: csharp/csharp_class_features/get_hash_code_override_consistent_for_same_data
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Key{int V;public Key(int v){V=v;}public override int GetHashCode()=>V.GetHashCode();}
__Check((new Key(7).GetHashCode()==new Key(7).GetHashCode()).ToString(), "True");
