//! Nested class/struct/enum access from outer scope and private nested visibility.


csharp_cases! {
    nested_access_outer_instantiates_public_nested_class => {
        r#"class Shell{public class Core{public int Id=7;}} Console.WriteLine(new Shell.Core().Id);"#,
        ["7"]
    };

    nested_access_outer_method_returns_nested_instance => {
        r#"class Factory{public class Item{public string Tag="x";} public Item Build()=>new Item();} Console.WriteLine(new Factory().Build().Tag);"#,
        ["x"]
    };

    nested_access_outer_static_method_creates_nested => {
        r#"class Hub{public class Node{public int V=3;} public static Node Create()=>new Node();} Console.WriteLine(Hub.Create().V);"#,
        ["3"]
    };

    nested_access_fully_qualified_name_from_outside => {
        r#"class A{public class B{public int N=11;}} Console.WriteLine(new A.B().N);"#,
        ["11"]
    };

    nested_access_private_nested_via_outer_public_wrapper => {
        r#"class Vault{class Key{public int Code=99;} public int Open()=>new Key().Code;} Console.WriteLine(new Vault().Open());"#,
        ["99"]
    };

    nested_access_private_nested_static_from_outer => {
        r#"class Cache{static class Store{public static int V=5;} public static int Read()=>Store.V;} Console.WriteLine(Cache.Read());"#,
        ["5"]
    };

    nested_access_nested_reads_outer_instance_field => {
        r#"class Outer{int seed=4; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Read()=>o.seed;} public int Via()=>new Inner(this).Read();} Console.WriteLine(new Outer().Via());"#,
        ["4"]
    };

    nested_access_nested_invokes_outer_instance_method => {
        r#"class Outer{int Double(int n)=>n*2; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Run(int n)=>o.Double(n);} public int Via(int n)=>new Inner(this).Run(n);} Console.WriteLine(new Outer().Via(6));"#,
        ["12"]
    };

    nested_access_nested_reads_outer_static_field => {
        r#"class Outer{static int tally=8; public class Inner{public int Read()=>tally;}} Console.WriteLine(new Outer.Inner().Read());"#,
        ["8"]
    };

    nested_access_nested_invokes_outer_static_method => {
        r#"class Outer{static int Triple(int n)=>n*3; public class Inner{public int Run(int n)=>Triple(n);}} Console.WriteLine(new Outer.Inner().Run(2));"#,
        ["6"]
    };

    nested_access_nested_struct_from_outer => {
        r#"class Map{public struct Point{public int X; public int Y;} public Point Origin()=>new Point{X=0,Y=0};} var p=new Map().Origin(); Console.WriteLine(p.X);"#,
        ["0"]
    };

    nested_access_nested_struct_field_mutation => {
        r#"class Canvas{public struct Dot{public int X;} public Dot Make(){var d=new Dot(); d.X=9; return d;}} Console.WriteLine(new Canvas().Make().X);"#,
        ["9"]
    };

    nested_access_nested_enum_from_outer_method => {
        r#"class Job{public enum State{Idle,Busy} public State Current()=>State.Busy;} Console.WriteLine(new Job().Current());"#,
        ["Busy"]
    };

    nested_access_nested_enum_int_cast_from_outside => {
        r#"class Job{public enum State{Idle=0,Busy=1}} Console.WriteLine((int)Job.State.Busy);"#,
        ["1"]
    };

    nested_access_nested_enum_switch_in_outer => {
        r#"class Gate{public enum Mode{On,Off} public string Label(Mode m){switch(m){case Mode.On:return "on"; default:return "off";}}} Console.WriteLine(new Gate().Label(Gate.Mode.On));"#,
        ["on"]
    };

    nested_access_two_nested_classes_same_outer => {
        r#"class Pair{public class Left{public int V=1;} public class Right{public int V=2;}} Console.WriteLine(new Pair.Left().V); Console.WriteLine(new Pair.Right().V);"#,
        ["1", "2"]
    };

    nested_access_sibling_nested_types_independent => {
        r#"class Duo{public class A{public int Bump(int n)=>n+1;} public class B{public int Bump(int n)=>n+2;}} Console.WriteLine(new Duo.A().Bump(5)); Console.WriteLine(new Duo.B().Bump(5));"#,
        ["6", "7"]
    };

    nested_access_deeply_nested_class_chain => {
        r#"class L1{public class L2{public class L3{public int V=42;}}} Console.WriteLine(new L1.L2.L3().V);"#,
        ["42"]
    };

    nested_access_nested_static_inside_nested_static => {
        r#"class Root{public static class A{public static class B{public static int V=13;}}} Console.WriteLine(Root.A.B.V);"#,
        ["13"]
    };

    nested_access_nested_class_implements_outer_interface => {
        r#"class Host{public interface IRun{int Go();} public class Worker:IRun{public int Go()=>4;}} Console.WriteLine(new Host.Worker().Go());"#,
        ["4"]
    };

    nested_access_nested_interface_implemented_by_nested_class => {
        r#"class Device{public interface IPort{string Open();} public class Usb:IPort{public string Open()=>"usb";}} Console.WriteLine(new Device.Usb().Open());"#,
        ["usb"]
    };

    nested_access_generic_outer_nested_uses_type_param => {
        r#"class Box<T>{public class Holder{public T Value;}} var h=new Box<int>.Holder(); h.Value=15; Console.WriteLine(h.Value);"#,
        ["15"]
    };

    nested_access_generic_outer_nested_string => {
        r#"class Box<T>{public class Holder{public T Value;} public Holder(T v){Value=v;}} Console.WriteLine(new Box<string>.Holder("ok").Value);"#,
        ["ok"]
    };

    nested_access_nested_class_inherits_nested_base => {
        r#"class Shapes{public class Base{public virtual string Name()=>"base";} public class Circle:Base{public override string Name()=>"circle";}} Console.WriteLine(new Shapes.Circle().Name());"#,
        ["circle"]
    };

    nested_access_outer_exposes_nested_via_property => {
        r#"class Shell{public class Core{public int Id=2;} Core _c=new Core(); public Core Inner=>_c;} Console.WriteLine(new Shell().Inner.Id);"#,
        ["2"]
    };

    nested_access_outer_field_holds_nested_struct => {
        r#"class Grid{public struct Cell{public int V;} Cell _c; public Grid(){_c.V=6;} public int Read()=>_c.V;} Console.WriteLine(new Grid().Read());"#,
        ["6"]
    };

    nested_access_private_nested_enum_via_public_method => {
        r#"class Status{enum Code{Ok=0,Fail=1} public int Read()=>(int)Code.Ok;} Console.WriteLine(new Status().Read());"#,
        ["0"]
    };

    nested_access_private_nested_struct_via_factory => {
        r#"class Builder{struct Part{public int N;} Part Make(){return new Part{N=8};} public int Build()=>Make().N;} Console.WriteLine(new Builder().Build());"#,
        ["8"]
    };

    nested_access_nested_delegate_declared_in_outer => {
        r#"class MathUtil{public delegate int Op(int a,int b); public class Calc{public int Run(Op f,int a,int b)=>f(a,b);}} Console.WriteLine(new MathUtil.Calc().Run((x,y)=>x+y,2,3));"#,
        ["5"]
    };

    nested_access_nested_class_static_factory_method => {
        r#"class Pool{public class Token{public int Id; public static Token Make(int id)=>new Token{Id=id};}} Console.WriteLine(Pool.Token.Make(21).Id);"#,
        ["21"]
    };

    nested_access_nested_enum_flags_in_outer => {
        r#"class Auth{[System.Flags] public enum Perm{None=0,Read=1,Write=2} public Perm All()=>Perm.Read|Perm.Write;} Console.WriteLine((int)new Auth().All());"#,
        ["3"]
    };

    nested_access_outer_static_nested_enum_member => {
        r#"class Config{public enum Level{Low,High} public static Level Default=>Level.Low;} Console.WriteLine(Config.Default);"#,
        ["Low"]
    };

    nested_access_nested_class_captures_outer_const => {
        r#"class Outer{public const string Prefix="pre"; public class Inner{public string Tag()=>Prefix+"fix";}} Console.WriteLine(new Outer.Inner().Tag());"#,
        ["prefix"]
    };

    nested_access_nested_class_captures_outer_readonly => {
        r#"class Outer{public readonly int Seed=10; public class Inner{Outer o; public Inner(Outer o){this.o=o;} public int Read()=>o.Seed;} public int Via()=>new Inner(this).Read();} Console.WriteLine(new Outer().Via());"#,
        ["10"]
    };

    nested_access_nested_in_partial_outer_part => {
        r#"partial class Worker{public class Helper{public int Run()=>1;}} partial class Worker{public int Go()=>new Helper().Run();} Console.WriteLine(new Worker().Go());"#,
        ["1"]
    };

    nested_access_nested_record_style_class => {
        r#"class Orders{public class Line{public int Qty; public int Total()=>Qty*2;} public Line Make(int q)=>new Line{Qty=q};} Console.WriteLine(new Orders().Make(4).Total());"#,
        ["8"]
    };

    nested_access_nested_visibility_public_from_external => {
        r#"class Api{public class Endpoint{public string Path="/v1";}} Console.WriteLine(new Api.Endpoint().Path);"#,
        ["/v1"]
    };

    nested_access_private_nested_not_exposed_but_outer_delegates => {
        r#"class Service{class Engine{public string Run()=>"ok";} public string Execute()=>new Engine().Run();} Console.WriteLine(new Service().Execute());"#,
        ["ok"]
    };

    nested_access_nested_struct_copy_independent => {
        r#"class Sheet{public struct Cell{public int V;} public int Sum(){var a=new Cell(); var b=a; a.V=3; b.V=5; return a.V+b.V;}} Console.WriteLine(new Sheet().Sum());"#,
        ["8"]
    };

    nested_access_nested_class_list_of_nested => {
        r#"class Bag{public class Item{public int Id;} public System.Collections.Generic.List<Item> All(){var list=new System.Collections.Generic.List<Item>(); list.Add(new Item{Id=1}); return list;}} Console.WriteLine(new Bag().All()[0].Id);"#,
        ["1"]
    };

    nested_access_nested_static_class_utility => {
        r#"class Text{public static class Util{public static string Join(string a,string b)=>a+b;} public static string Merge()=>Util.Join("a","b");} Console.WriteLine(Text.Merge());"#,
        ["ab"]
    };

    nested_access_nested_type_name_via_gettype => {
        r#"class Outer{public class Inner{}} Console.WriteLine(typeof(Outer.Inner).Name);"#,
        ["Inner"]
    };

    nested_access_outer_passes_nested_to_method => {
        r#"class Store{public class Item{public int Id;} int Inspect(Item i)=>i.Id; public int Check()=>Inspect(new Item{Id=44});} Console.WriteLine(new Store().Check());"#,
        ["44"]
    };

    nested_access_nested_enum_to_string => {
        r#"class Mode{public enum Kind{Alpha,Beta} public string Label()=>Kind.Beta.ToString();} Console.WriteLine(new Mode().Label());"#,
        ["Beta"]
    };

    nested_access_nested_class_with_outer_generic_constraint => {
        r#"class Repo<T>{public class Row{public T Data;} public T Read(Row r)=>r.Data;} Console.WriteLine(new Repo<int>().Read(new Repo<int>.Row{Data=77}));"#,
        ["77"]
    };
}
