// vybe-test: csharp/csharp_initialization_order/static_constructor_runs_once_before_first_instance_creation
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

class Counter {
    static Counter() { Console.WriteLine("static-ctor"); }
    public Counter() { Console.WriteLine("instance"); }
}
new Counter();
new Counter();
