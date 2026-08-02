// vybe-test: csharp/csharp_custom_event_accessors/custom_event_lambda_capture
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; var b=new Btn(); b.Click+=()=>{n++; Console.WriteLine(n);}; b.Raise(); b.Raise();
