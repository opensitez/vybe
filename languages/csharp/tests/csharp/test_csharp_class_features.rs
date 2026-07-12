//! Class features: static members, object clone, `MemberwiseClone`, finalisers.
use super::helpers::run_csharp;

#[test]
fn static_field_shared_across_all_instances() {
    assert_eq!(
        run_csharp(
            r#"class Ctr{public static int Count=0; public Ctr(){Count++;}}
new Ctr(); new Ctr(); new Ctr();
Console.WriteLine(Ctr.Count);"#
        ),
        &["3"]
    );
}

#[test]
fn memberwise_clone_creates_shallow_copy() {
    assert_eq!(
        run_csharp(
            r#"class Point:System.ICloneable{public int X,Y;public object Clone()=>MemberwiseClone();}
var a=new Point{X=1,Y=2};
var b=(Point)a.Clone();
b.X=99;
Console.WriteLine(a.X);"#
        ),
        &["1"]
    );
}

#[test]
fn static_method_returns_instance_via_factory() {
    assert_eq!(
        run_csharp(
            r#"class Logger{
    string prefix;
    Logger(string p){prefix=p;}
    public static Logger For(string name)=>new Logger(name);
    public string Format(string m)=>$"[{prefix}] {m}";
}
Console.WriteLine(Logger.For("app").Format("hello"));"#
        ),
        &["[app] hello"]
    );
}

#[test]
fn object_to_string_default_returns_type_name() {
    assert_eq!(
        run_csharp(
            r#"class Widget{}
Console.WriteLine(new Widget().ToString());"#
        ),
        &["Widget"]
    );
}

#[test]
fn override_to_string_produces_custom_output() {
    assert_eq!(
        run_csharp(
            r#"class Color{int R,G,B;public Color(int r,int g,int b){R=r;G=g;B=b;}
public override string ToString()=>$"rgb({R},{G},{B})";}
Console.WriteLine(new Color(255,0,128));"#
        ),
        &["rgb(255,0,128)"]
    );
}

#[test]
fn get_hash_code_override_consistent_for_same_data() {
    assert_eq!(
        run_csharp(
            r#"class Key{int V;public Key(int v){V=v;}public override int GetHashCode()=>V.GetHashCode();}
Console.WriteLine(new Key(7).GetHashCode()==new Key(7).GetHashCode());"#
        ),
        &["True"]
    );
}
