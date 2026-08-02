// vybe-test: csharp/csharp_error_handling/throw_from_method
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

int Divide(int a, int b) {
    if (b == 0) throw new Exception("Division by zero");
    return a / b;
}
try {
    Console.WriteLine(Divide(10, 2));
    Console.WriteLine(Divide(10, 0));
} catch (Exception e) {
    Console.WriteLine(e.Message);
}
