// vybe-test: csharp/csharp_static_constructor_once/static_constructor_increments_counter_only_once_across_two_instance_allocations
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_once.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Tracker {
    public static int Instances;
    static Tracker() { Instances++; }
}
_ = new Tracker();
_ = new Tracker();
__Check((Tracker.Instances).ToString(), "1");
