// vybe-test: csharp/csharp_delegates/action_delegate
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

Action<string> greet = name => Console.WriteLine("Hello " + name);
greet("World");
greet("Alice");
