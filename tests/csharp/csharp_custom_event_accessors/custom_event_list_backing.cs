// vybe-test: csharp/csharp_custom_event_accessors/custom_event_list_backing
// origin: languages/csharp/tests/csharp/test_csharp_custom_event_accessors.rs

class Btn{var _list=new System.Collections.Generic.List<System.Action>(); public event System.Action Click{add{_list.Add(value);} remove{_list.Remove(value);}} public void Raise(){foreach(var h in _list) h();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Click+=()=>n+=2; b.Raise(); Console.WriteLine(n);
