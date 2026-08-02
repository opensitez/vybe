// vybe-test: csharp/csharp_class_features/static_field_shared_across_all_instances
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Ctr{public static int Count=0; public Ctr(){Count++;}}
new Ctr(); new Ctr(); new Ctr();
__Check((Ctr.Count).ToString(), "3");
