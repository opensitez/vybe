//! Advanced property patterns: computed, lazy-init, change notification, indexer-backed.
use super::helpers::run_csharp;

#[test]
fn computed_property_recalculates_on_each_access() {
    assert_eq!(
        run_csharp(
            r#"class Circle{
    public double Radius;
    public double Circumference=>2*System.Math.PI*Radius;
}
var c=new Circle{Radius=1.0};
Console.WriteLine(System.Math.Round(c.Circumference,5));"#
        ),
        &["6.28319"]
    );
}

#[test]
fn lazy_initialized_property_created_on_first_access() {
    assert_eq!(
        run_csharp(
            r#"class Config{
    System.Lazy<string> _tag=new System.Lazy<string>(()=>"computed");
    public string Tag=>_tag.Value;
}
var c=new Config();
Console.WriteLine(c.Tag);"#
        ),
        &["computed"]
    );
}

#[test]
fn property_with_backing_field_validation() {
    assert_eq!(
        run_csharp(
            r#"class Age{
    int _value;
    public int Value{
        get=>_value;
        set{if(value<0)throw new System.ArgumentException();_value=value;}
    }
}
var a=new Age{Value=25};
Console.WriteLine(a.Value);"#
        ),
        &["25"]
    );
}

#[test]
fn static_property_shared_across_instances() {
    assert_eq!(
        run_csharp(
            r#"class AppConfig{public static string Version{get;set;}="1.0";}
AppConfig.Version="2.0";
Console.WriteLine(new System.Object().GetType()!=null);
Console.WriteLine(AppConfig.Version);"#
        ),
        &["True", "2.0"]
    );
}

#[test]
fn property_change_fires_on_setter_invocation() {
    assert_eq!(
        run_csharp(
            r#"class Observable:System.ComponentModel.INotifyPropertyChanged{
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
Console.WriteLine(notified);"#
        ),
        &["True"]
    );
}
