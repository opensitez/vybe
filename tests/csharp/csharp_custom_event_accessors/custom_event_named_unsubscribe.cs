// vybe-test: csharp/csharp_custom_event_accessors/custom_event_named_unsubscribe
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} void OnClick(){Console.WriteLine("hit");} var b=new Btn(); b.Click+=OnClick; b.Click-=OnClick; b.Raise(); Console.WriteLine("done");
