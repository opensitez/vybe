// vybe-test: csharp/csharp_properties_advanced/property_change_fires_on_setter_invocation
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Observable:System.ComponentModel.INotifyPropertyChanged{
    public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
    string _name="";
    public string Name{
        get=>_name;
        set{_name=value;PropertyChanged?.Invoke(this,new System.ComponentModel.PropertyChangedEventArgs(nameof(Name)));}
    }
}
var o=new Observable();
bool notified=false;
o.PropertyChanged+=(_,__)=>notified=true;
o.Name="Alice";
__Check((notified).ToString(), "True");
