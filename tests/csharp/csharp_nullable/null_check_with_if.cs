// vybe-test: csharp/csharp_nullable/null_check_with_if
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

string s = null;
if (s == null) {
    Console.WriteLine("is null");
} else {
    Console.WriteLine("has value");
}
s = "test";
if (s != null) {
    Console.WriteLine("has value");
}
