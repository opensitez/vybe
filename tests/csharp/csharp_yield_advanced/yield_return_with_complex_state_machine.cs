// vybe-test: csharp/csharp_yield_advanced/yield_return_with_complex_state_machine
// origin: languages/csharp/tests/csharp/test_csharp_yield_advanced.rs

System.Collections.Generic.IEnumerable<string> Words(string s){
    var parts=s.Split(' ');
    foreach(var p in parts) if(p.Length>0) yield return p;
}
Console.WriteLine(string.Join("|",Words("hello  world  foo")));
