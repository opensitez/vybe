//! Default interface method implementations — deep coverage including diamond resolution via implementing classes.

csharp_cases! {
    default_interface_method_called_on_concrete_without_override => {
        r#"interface IGreet{string Hello()=>"hi";} class Person:IGreet{} Console.WriteLine(new Person().Hello());"#,
        ["hi"]
    };

    default_interface_method_visible_through_interface_reference => {
        r#"interface ICalc{int Double(int n)=>n*2;} class Worker:ICalc{} ICalc w=new Worker(); Console.WriteLine(w.Double(4));"#,
        ["8"]
    };

    default_interface_method_override_in_class => {
        r#"interface IFormat{string Show(int n)=>n.ToString();} class Custom:IFormat{public string Show(int n)=>"x"+n;} Console.WriteLine(new Custom().Show(3));"#,
        ["x3"]
    };

    default_interface_method_class_override_via_interface_typed => {
        r#"interface IFormat{string Show(int n)=>n.ToString();} class Custom:IFormat{public string Show(int n)=>"x"+n;} IFormat f=new Custom(); Console.WriteLine(f.Show(3));"#,
        ["x3"]
    };

    diamond_two_defaults_resolved_by_class_public_override => {
        r#"interface IA{void M()=>Console.WriteLine("A");} interface IB{void M()=>Console.WriteLine("B");} class C:IA,IB{public void M()=>Console.WriteLine("C");} new C().M();"#,
        ["C"]
    };

    diamond_two_defaults_resolved_by_explicit_interface_impl => {
        r#"interface IA{void M()=>Console.WriteLine("A");} interface IB{void M()=>Console.WriteLine("B");} class C:IA,IB{void IA.M()=>Console.WriteLine("IA"); void IB.M()=>Console.WriteLine("IB");} ((IA)new C()).M(); ((IB)new C()).M();"#,
        ["IA", "IB"]
    };

    default_method_calls_other_interface_method => {
        r#"interface IBase{string Core()=>"core"; string Wrap()=>"["+Core()+"]";} class Node:IBase{} Console.WriteLine(new Node().Wrap());"#,
        ["[core]"]
    };

    default_method_uses_instance_property => {
        r#"interface IHas{int N{get;} int Twice()=>N*2;} class Box:IHas{public int N{get;set;}} var b=new Box{N=5}; Console.WriteLine(b.Twice());"#,
        ["10"]
    };

    two_interfaces_different_default_methods_both_callable => {
        r#"interface IA{int A()=>1;} interface IB{int B()=>2;} class Both:IA,IB{} var x=new Both(); Console.WriteLine(x.A()+x.B());"#,
        ["3"]
    };

    default_method_on_generic_interface => {
        r#"interface IBox<T>{T Echo(T v)=>v;} class IntBox:IBox<int>{} Console.WriteLine(new IntBox().Echo(7));"#,
        ["7"]
    };

    default_method_chain_three_levels => {
        r#"interface I1{string S()=>"a";} interface I2:I1{string T()=>S()+"b";} class X:I2{} Console.WriteLine(new X().T());"#,
        ["ab"]
    };

    default_method_replaced_only_on_one_interface_in_diamond => {
        r#"interface IA{string Tag()=>"A";} interface IB{string Tag()=>"B";} class Pick:IA,IB{public string Tag()=>"P";} Console.WriteLine(new Pick().Tag());"#,
        ["P"]
    };

    default_void_method_with_side_effect => {
        r#"interface ILog{void Ping(){Console.WriteLine("ping");}} class Silent:ILog{} new Silent().Ping();"#,
        ["ping"]
    };

    default_method_returning_bool => {
        r#"interface ICheck{bool Ok()=>true;} class Gate:ICheck{} Console.WriteLine(new Gate().Ok());"#,
        ["True"]
    };

    default_method_with_parameters => {
        r#"interface IAdd{int Sum(int a,int b)=>a+b;} class Calc:IAdd{} Console.WriteLine(new Calc().Sum(2,5));"#,
        ["7"]
    };

    default_method_string_interpolation => {
        r#"interface IName{string Name{get;} string Label()=>$"name={Name}";} class User:IName{public string Name{get;set;}="Ann";} Console.WriteLine(new User().Label());"#,
        ["name=Ann"]
    };

    default_method_struct_implementor => {
        r#"interface IArea{double Area()=>1.0;} struct Unit:IArea{} Console.WriteLine(new Unit().Area());"#,
        ["1"]
    };

    default_method_not_visible_as_class_member_without_interface => {
        r#"interface IHidden{int Secret()=>9;} class Worker:IHidden{} IHidden w=new Worker(); Console.WriteLine(w.Secret());"#,
        ["9"]
    };

    default_method_override_calls_base_default_via_super_not_available_use_explicit => {
        r#"interface IA{string V()=>"a";} interface IB:IA{string W()=>V()+"b";} class Z:IB{public string V()=>"z";} Console.WriteLine(((IB)new Z()).W());"#,
        ["ab"]
    };

    default_interface_property_getter => {
        r#"interface IProp{int Max=>100;} class Reader:IProp{} Console.WriteLine(new Reader().Max);"#,
        ["100"]
    };

    default_interface_property_setter_pair => {
        r#"interface ICounter{int Count{get;set;} void Inc(){Count++;}} class C:ICounter{public int Count{get;set;}} var c=new C(); c.Inc(); Console.WriteLine(c.Count);"#,
        ["1"]
    };

    diamond_three_interfaces_class_unified_override => {
        r#"interface IA{void P()=>Console.WriteLine("A");} interface IB{void P()=>Console.WriteLine("B");} interface IC{void P()=>Console.WriteLine("C");} class U:IA,IB,IC{public void P()=>Console.WriteLine("U");} new U().P();"#,
        ["U"]
    };

    default_method_invokes_virtual_class_method => {
        r#"interface IRun{string Go(){return Run();} string Run();} class Job:IRun{public string Run()=>"done";} Console.WriteLine(new Job().Go());"#,
        ["done"]
    };

    default_method_on_interface_with_multiple_members => {
        r#"interface IOps{int Add(int a,int b)=>a+b; int Mul(int a,int b)=>a*b;} class Ops:IOps{} var o=new Ops(); Console.WriteLine(o.Add(2,3)+o.Mul(2,3));"#,
        ["11"]
    };

    default_method_reused_by_two_classes => {
        r#"interface IDouble{int Twice(int n)=>n*2;} class A:IDouble{} class B:IDouble{} Console.WriteLine(new A().Twice(3)+new B().Twice(4));"#,
        ["14"]
    };

    default_method_one_class_overrides_other_uses_default => {
        r#"interface IScale{int Scale(int n)=>n;} class Plain:IScale{} class Double:IScale{public int Scale(int n)=>n*2;} Console.WriteLine(new Plain().Scale(5)+new Double().Scale(5));"#,
        ["15"]
    };

    default_method_with_null_conditional => {
        r#"interface IMaybe{string Name{get;} string Safe()=>Name??"none";} class Item:IMaybe{public string Name{get;set;}} Console.WriteLine(new Item().Safe());"#,
        ["none"]
    };

    default_method_bool_logic => {
        r#"interface IFlag{bool On{get;} bool IsOff()=>!On;} class Switch:IFlag{public bool On{get;set;}=true;} Console.WriteLine(new Switch().IsOff());"#,
        ["False"]
    };

    default_method_char_conversion => {
        r#"interface IChar{char C{get;} string AsString()=>C.ToString();} class Letter:IChar{public char C{get;set;}='Q';} Console.WriteLine(new Letter().AsString());"#,
        ["Q"]
    };

    default_method_decimal_math => {
        r#"interface IMoney{decimal Add(decimal a,decimal b)=>a+b;} class Wallet:IMoney{} Console.WriteLine(new Wallet().Add(1.5m,2.5m));"#,
        ["4.0"]
    };

    default_method_list_count_helper => {
        r#"interface ISize{int Len(System.Collections.Generic.List<int> xs)=>xs.Count;} class Measurer:ISize{} Console.WriteLine(new Measurer().Len(new System.Collections.Generic.List<int>{1,2,3}));"#,
        ["3"]
    };

    default_method_explicit_interface_route_avoids_diamond => {
        r#"interface IA{void Show()=>Console.WriteLine("A");} interface IB{void Show()=>Console.WriteLine("B");} class Split:IA,IB{void IA.Show()=>Console.WriteLine("IA"); void IB.Show()=>Console.WriteLine("IB");} ((IA)new Split()).Show();"#,
        ["IA"]
    };

    default_method_inherited_through_subinterface => {
        r#"interface IRoot{int Base()=>1;} interface IChild:IRoot{int Child()=>Base()+1;} class Node:IChild{} Console.WriteLine(new Node().Child());"#,
        ["2"]
    };

    default_method_override_in_subclass_of_implementor => {
        r#"interface IVal{int Get()=>0;} class Base:IVal{} class Derived:Base,IVal{public int Get()=>5;} Console.WriteLine(new Derived().Get());"#,
        ["5"]
    };

    default_method_accesses_static_class_field => {
        r#"interface IStatic{int Read()=>Holder.N;} static class Holder{public static int N=8;} class R:IStatic{} Console.WriteLine(new R().Read());"#,
        ["8"]
    };

    default_method_string_builder_pattern => {
        r#"interface IBuild{string Step1()=>"a"; string Step2()=>Step1()+"b";} class Chain:IBuild{} Console.WriteLine(new Chain().Step2());"#,
        ["ab"]
    };

    diamond_resolution_prefers_most_specific_class_method => {
        r#"interface IA{string Tag()=>"A";} interface IB:IA{string Tag()=>"B";} class Leaf:IB{public string Tag()=>"L";} Console.WriteLine(new Leaf().Tag());"#,
        ["L"]
    };

    default_method_with_enum_parameter => {
        r#"enum Mode{On,Off} interface IMode{string Label(Mode m)=>m.ToString();} class M:IMode{} Console.WriteLine(new M().Label(Mode.On));"#,
        ["On"]
    };

    default_method_indexer_helper => {
        r#"interface IIdx{string this[int i]{get;} string First()=>this[0];} class Arr:IIdx{public string this[int i]=>"v"+i;} Console.WriteLine(new Arr().First());"#,
        ["v0"]
    };

    default_method_two_defaults_only_one_overridden => {
        r#"interface IA{int A()=>1;} interface IB{int B()=>2;} class Mix:IA,IB{public int A()=>10;} var m=new Mix(); Console.WriteLine(m.A()+m.B());"#,
        ["12"]
    };

    default_method_recursive_default_calls_itself => {
        r#"interface IRec{int Fact(int n)=>n<=1?1:n*Fact(n-1);} class Math:IRec{} Console.WriteLine(new Math().Fact(5));"#,
        ["120"]
    };

    default_method_nullable_int_handling => {
        r#"interface INull{int? N{get;} int OrZero()=>N??0;} class Maybe:INull{public int? N{get;set;}=null;} Console.WriteLine(new Maybe().OrZero());"#,
        ["0"]
    };

    default_method_object_to_string_fallback => {
        r#"interface IObj{string Desc()=>ToString();} class Thing:IObj{public override string ToString()=>"thing";} Console.WriteLine(new Thing().Desc());"#,
        ["thing"]
    };

    default_method_interface_hiding_with_new_class_method => {
        r#"interface IHide{int V()=>1;} class Hide:IHide{public new int V()=>2;} Console.WriteLine(new Hide().V());"#,
        ["2"]
    };

    default_method_multiple_interface_inheritance_same_signature => {
        r#"interface IA{int Score()=>1;} interface IB{int Score()=>2;} class Dual:IA,IB{public int Score()=>3;} Console.WriteLine(new Dual().Score());"#,
        ["3"]
    };

    default_method_void_chain_two_defaults => {
        r#"interface IA{void A(){Console.WriteLine("a");}} interface IB{void B(){Console.WriteLine("b");}} class Both:IA,IB{} var b=new Both(); b.A(); b.B();"#,
        ["a", "b"]
    };

    default_method_generic_constraint_on_interface => {
        r#"interface ICompare<T> where T:System.IComparable<T>{int Cmp(T a,T b)=>a.CompareTo(b);} class S:ICompare<int>{} Console.WriteLine(new S().Cmp(3,7));"#,
        ["-1"]
    };

    default_method_with_local_variable => {
        r#"interface ILocal{int Triple(int n){var t=n*3; return t;}} class L:ILocal{} Console.WriteLine(new L().Triple(4));"#,
        ["12"]
    };

    default_method_switch_expression_body => {
        r#"interface ISw{string Code(int n)=>n switch{1=>"one",2=>"two",_=>"many"};} class C:ISw{} Console.WriteLine(new C().Code(2));"#,
        ["two"]
    };

    default_method_diamond_class_picks_single_public_impl => {
        r#"interface IA{void Print()=>Console.WriteLine("A");} interface IB{void Print()=>Console.WriteLine("B");} class Merge:IA,IB{public void Print()=>Console.WriteLine("M");} new Merge().Print();"#,
        ["M"]
    };
}
