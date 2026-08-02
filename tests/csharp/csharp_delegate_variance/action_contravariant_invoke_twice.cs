// vybe-test: csharp/csharp_delegate_variance/action_contravariant_invoke_twice
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

System.Action<object> w=v=>Console.WriteLine(v); System.Action<string> n=w; n("a"); n("b");
