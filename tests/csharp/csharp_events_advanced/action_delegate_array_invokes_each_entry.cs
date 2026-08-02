// vybe-test: csharp/csharp_events_advanced/action_delegate_array_invokes_each_entry
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using System; Action[] actions = { () => Console.WriteLine("one"), () => Console.WriteLine("two") }; foreach (var action in actions) action();
