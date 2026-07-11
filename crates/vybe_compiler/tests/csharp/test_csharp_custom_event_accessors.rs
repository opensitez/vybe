//! Custom event add/remove accessors with backing fields and explicit raise methods.

csharp_cases! {
    custom_event_invoke_via_backing => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_remove_stops_invoke => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; System.Action h=()=>n++; var b=new Btn(); b.Click+=h; b.Click-=h; b.Raise(); Console.WriteLine(n);"#,
        ["0"]
    };

    custom_event_two_handlers => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}var log=""; var b=new Btn(); b.Click+=()=>log+="a"; b.Click+=()=>log+="b"; b.Raise(); Console.WriteLine(log);"#,
        ["ab"]
    };

    custom_event_remove_one_of_two => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}var log=""; System.Action h=()=>log+="a"; var b=new Btn(); b.Click+=h; b.Click+=()=>log+="b"; b.Click-=h; b.Raise(); Console.WriteLine(log);"#,
        ["b"]
    };

    custom_event_null_backing_safe => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}var b=new Btn(); b.Raise(); Console.WriteLine("ok");"#,
        ["ok"]
    };

    custom_event_resubscribe_after_remove => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; System.Action h=()=>n++; var b=new Btn(); b.Click+=h; b.Click-=h; b.Click+=h; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_same_handler_twice => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; System.Action h=()=>n++; var b=new Btn(); b.Click+=h; b.Click+=h; b.Raise(); Console.WriteLine(n);"#,
        ["2"]
    };

    custom_event_remove_unsubscribed => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}System.Action h=()=>{}; var b=new Btn(); b.Click-=h; b.Raise(); Console.WriteLine("fine");"#,
        ["fine"]
    };

    custom_event_count_tracked => {
        r#"class Btn{System.Action _e; int _count; public event System.Action Tick{add{_e+=value;_count++;} remove{_e-=value;_count--;}} public int Count=>_count; public void Fire(){_e?.Invoke();}} var b=new Btn(); System.Action h=()=>{}; b.Tick+=h; b.Tick+=()=>{}; b.Tick-=h; Console.WriteLine(b.Count);"#,
        ["1"]
    };

    custom_event_eventhandler_backing => {
        r#"class Src{System.EventHandler _h; public event System.EventHandler Changed{add{_h+=value;} remove{_h-=value;}} public void Raise(){_h?.Invoke(this,System.EventArgs.Empty);}} int n=0; var s=new Src(); s.Changed+=(o,e)=>n++; s.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_eventhandler_generic => {
        r#"class Msg: System.EventArgs{public string Text;} class Ch{System.EventHandler<Msg> _h; public event System.EventHandler<Msg> Sent{add{_h+=value;} remove{_h-=value;}} public void Emit(string t){_h?.Invoke(this,new Msg{Text=t});}} string out_=""; var c=new Ch(); c.Sent+=(o,e)=>out_=e.Text; c.Emit("hi"); Console.WriteLine(out_);"#,
        ["hi"]
    };

    custom_event_lambda_capture => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; var b=new Btn(); b.Click+=()=>{n++; Console.WriteLine(n);}; b.Raise(); b.Raise();"#,
        ["1", "2"]
    };

    custom_event_named_method => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} void OnClick(){Console.WriteLine("hit");} var b=new Btn(); b.Click+=OnClick; b.Raise();"#,
        ["hit"]
    };

    custom_event_named_unsubscribe => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} void OnClick(){Console.WriteLine("hit");} var b=new Btn(); b.Click+=OnClick; b.Click-=OnClick; b.Raise(); Console.WriteLine("done");"#,
        ["done"]
    };

    custom_event_two_instances_independent => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int a=0,b=0; var x=new Btn(); var y=new Btn(); x.Click+=()=>a++; y.Click+=()=>b++; x.Raise(); Console.WriteLine(a); Console.WriteLine(b);"#,
        ["1", "0"]
    };

    custom_event_wrong_instance_unsub => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; System.Action h=()=>n++; var a=new Btn(); var b=new Btn(); a.Click+=h; b.Click-=h; a.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_add_in_constructor => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public Btn(){_c+=()=>_boot=1;} int _boot; public void Raise(){_c?.Invoke();} public int Boot=>_boot;} var b=new Btn(); Console.WriteLine(b.Boot);"#,
        ["1"]
    };

    custom_event_prevent_duplicate => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{if(_c==null||!_c.GetInvocationList().Contains(value)) _c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; System.Action h=()=>n++; var b=new Btn(); b.Click+=h; b.Click+=h; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_action_string => {
        r#"class Line{System.Action<string> _h; public event System.Action<string> Write{add{_h+=value;} remove{_h-=value;}} public void Emit(string s){_h?.Invoke(s);}} string log=""; var l=new Line(); l.Write+=s=>log+=s; l.Emit("x"); Console.WriteLine(log);"#,
        ["x"]
    };

    custom_event_multiple_raises => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); b.Raise(); b.Raise(); Console.WriteLine(n);"#,
        ["3"]
    };

    custom_event_clear_handlers => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();} public int Count=>_c==null?0:_c.GetInvocationList().Length;} System.Action a=()=>{}; System.Action b=()=>{}; var btn=new Btn(); btn.Click+=a; btn.Click+=b; btn.Click-=a; btn.Click-=b; Console.WriteLine(btn.Count);"#,
        ["0"]
    };

    custom_event_base_backing => {
        r#"class Base{System.Action _e; public event System.Action Ping{add{_e+=value;} remove{_e-=value;}} protected void OnPing(){_e?.Invoke();}} class Child:Base{public void Fire(){OnPing();}} int n=0; var c=new Child(); c.Ping+=()=>n++; c.Fire(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_derived_subscribes_base => {
        r#"class Base{System.Action _e; public event System.Action Ping{add{_e+=value;} remove{_e-=value;}} protected void OnPing(){_e?.Invoke();}} class Child:Base{public void Fire(){OnPing();}} int n=0; Child c=new Child(); c.Ping+=()=>n++; c.Fire(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_public_subscribe_private_raise => {
        r#"class Hub{System.Action _h; public event System.Action Signal{add{_h+=value;} remove{_h-=value;}} public void Pulse(){_h?.Invoke();}} int v=0; var h=new Hub(); h.Signal+=()=>v=9; h.Pulse(); Console.WriteLine(v);"#,
        ["9"]
    };

    custom_event_non_null_add_guard => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{if(value!=null)_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_lock_in_accessor => {
        r#"class Btn{System.Action _c; readonly object _gate=new object(); public event System.Action Click{add{lock(_gate){_c+=value;}} remove{lock(_gate){_c-=value;}}} public void Raise(){_c?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_log_add_remove => {
        r#"class Btn{System.Action _c; public string Log=""; public event System.Action Click{add{Log+="+"; _c+=value;} remove{Log+="-"; _c-=value;}} public void Raise(){_c?.Invoke();}} System.Action h=()=>{}; var b=new Btn(); b.Click+=h; b.Click-=h; Console.WriteLine(b.Log);"#,
        ["+-"]
    };

    custom_event_handler_sets_state => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int state=0; var b=new Btn(); b.Click+=()=>state=42; b.Raise(); Console.WriteLine(state);"#,
        ["42"]
    };

    custom_event_subscribe_after_first_raise => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}var b=new Btn(); b.Raise(); b.Click+=()=>Console.WriteLine("late"); b.Raise();"#,
        ["late"]
    };

    custom_event_multicast_sum => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; System.Action h=()=>n++; var b=new Btn(); b.Click+=h; b.Click+=()=>n+=10; b.Raise(); Console.WriteLine(n);"#,
        ["11"]
    };

    custom_event_list_backing => {
        r#"class Btn{var _list=new System.Collections.Generic.List<System.Action>(); public event System.Action Click{add{_list.Add(value);} remove{_list.Remove(value);}} public void Raise(){foreach(var h in _list) h();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Click+=()=>n+=2; b.Raise(); Console.WriteLine(n);"#,
        ["3"]
    };

    custom_event_validate_non_null_add => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{if(value==null) return; _c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_has_after_add => {
        r#"class Btn{System.Action _c=null; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public bool Has=>_c!=null; public void Raise(){_c?.Invoke();}} var b=new Btn(); b.Click+=()=>{}; Console.WriteLine(b.Has);"#,
        ["True"]
    };

    custom_event_has_after_remove => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public bool Has=>_c!=null; public void Raise(){_c?.Invoke();}} System.Action h=()=>{}; var b=new Btn(); b.Click+=h; b.Click-=h; Console.WriteLine(b.Has);"#,
        ["False"]
    };

    custom_event_order_preserved => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} var log=new System.Collections.Generic.List<string>(); var b=new Btn(); b.Click+=()=>log.Add("1"); b.Click+=()=>log.Add("2"); b.Raise(); Console.WriteLine(string.Join(",",log));"#,
        ["1,2"]
    };

    custom_event_zero_after_clear => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; System.Action h=()=>n++; var b=new Btn(); b.Click+=h; b.Click-=h; b.Raise(); Console.WriteLine(n);"#,
        ["0"]
    };

    custom_event_action_int => {
        r#"class Src{System.Action<int> _h; public event System.Action<int> Value{add{_h+=value;} remove{_h-=value;}} public void Set(int v){_h?.Invoke(v);}} int got=0; var s=new Src(); s.Value+=v=>got=v; s.Set(15); Console.WriteLine(got);"#,
        ["15"]
    };

    custom_event_add_remove_counters => {
        r#"class Btn{System.Action _c; public int Adds=0; public int Removes=0; public event System.Action Click{add{Adds++; _c+=value;} remove{Removes++; _c-=value;}} public void Raise(){_c?.Invoke();}} System.Action h=()=>{}; var b=new Btn(); b.Click+=h; b.Click-=h; Console.WriteLine(b.Adds); Console.WriteLine(b.Removes);"#,
        ["1", "1"]
    };

    custom_event_not_called_before_subscribe => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}}int n=0; var b=new Btn(); b.Raise(); b.Click+=()=>n++; Console.WriteLine(n);"#,
        ["0"]
    };

    custom_event_clear_backing_null => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Clear(){_c=null;} public bool Empty=>_c==null;} var b=new Btn(); b.Click+=()=>{}; b.Clear(); Console.WriteLine(b.Empty);"#,
        ["True"]
    };

    custom_event_count_zero_no_subscribers => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Count=>_c==null?0:_c.GetInvocationList().Length; public void Raise(){_c?.Invoke();}} var b=new Btn(); b.Raise(); Console.WriteLine(b.Count);"#,
        ["0"]
    };

    custom_event_add_two_remove_one => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Count=>_c==null?0:_c.GetInvocationList().Length;} System.Action a=()=>{}; System.Action b=()=>{}; var btn=new Btn(); btn.Click+=a; btn.Click+=b; btn.Click-=a; Console.WriteLine(btn.Count);"#,
        ["1"]
    };

    custom_event_static_local_handler => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public void Raise(){_c?.Invoke();}} int n=0; void Bump(){n++;} var b=new Btn(); b.Click+=Bump; b.Raise(); Console.WriteLine(n);"#,
        ["1"]
    };

    custom_event_sender_passed_to_handler => {
        r#"class Src{System.EventHandler _h; public event System.EventHandler Tick{add{_h+=value;} remove{_h-=value;}} public void Pulse(){_h?.Invoke(this,System.EventArgs.Empty);}} object who=null; var s=new Src(); s.Tick+=(sender,e)=>who=sender; s.Pulse(); Console.WriteLine(who==s);"#,
        ["True"]
    };

    custom_event_subscriber_count_property => {
        r#"class Btn{System.Action _c; public event System.Action Click{add{_c+=value;} remove{_c-=value;}} public int Subscribers=>_c==null?0:_c.GetInvocationList().Length;} var b=new Btn(); b.Click+=()=>{}; b.Click+=()=>{}; Console.WriteLine(b.Subscribers);"#,
        ["2"]
    };
}
