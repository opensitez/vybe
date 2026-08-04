// vybe-test: csharp/csharp_properties_advanced/property_change_fires_on_setter_invocation
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((notified).ToString());
__Check("True");
