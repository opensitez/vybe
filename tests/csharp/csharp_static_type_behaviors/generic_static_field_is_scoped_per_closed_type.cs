// vybe-test: csharp/csharp_static_type_behaviors/generic_static_field_is_scoped_per_closed_type
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Cache<T> {
    public static int Hits;
}
Cache<int>.Hits++;
Cache<int>.Hits++;
Cache<string>.Hits++;
__Check((Cache<int>.Hits).ToString(), "2");
__Check((Cache<string>.Hits).ToString(), "1");
