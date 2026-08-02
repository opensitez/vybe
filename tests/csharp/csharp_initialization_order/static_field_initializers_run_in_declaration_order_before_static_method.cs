// vybe-test: csharp/csharp_initialization_order/static_field_initializers_run_in_declaration_order_before_static_method
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

class Logger {
    static string First = Mark("first");
    static string Second = Mark("second");
    static string Mark(string name) {
        Console.WriteLine(name);
        return name;
    }
    public static void Run() {
        Console.WriteLine("run");
    }
}
Logger.Run();
