//! Attribute usage patterns: Obsolete still callable, Flags enum, Serializable marker, Conditional via prints.


csharp_cases! {
    attribute_obsolete_method_still_invokes => {
        r#"using System; class S{[Obsolete("old")] public string Run()=>"ok";} Console.WriteLine(new S().Run());"#,
        ["ok"]
    };

    attribute_obsolete_void_method_still_runs => {
        r#"using System; class S{[Obsolete("legacy")] public void Ping(){Console.WriteLine("ping");}} new S().Ping();"#,
        ["ping"]
    };

    attribute_obsolete_property_getter_still_reads => {
        r#"using System; class S{[Obsolete("old")] public int Value=>42;} Console.WriteLine(new S().Value);"#,
        ["42"]
    };

    attribute_obsolete_on_class_instance_still_usable => {
        r#"using System; [Obsolete("old")] class S{public int N=1;} Console.WriteLine(new S().N);"#,
        ["1"]
    };

    attribute_obsolete_overload_both_callable => {
        r#"using System; class S{[Obsolete("a")] public int Go()=>1; [Obsolete("b")] public int Go(int x)=>x;} Console.WriteLine(new S().Go()); Console.WriteLine(new S().Go(5));"#,
        ["1", "5"]
    };

    attribute_obsolete_static_method_still_callable => {
        r#"using System; class S{[Obsolete("old")] public static int Id()=>9;} Console.WriteLine(S.Id());"#,
        ["9"]
    };

    attribute_obsolete_does_not_block_chained_calls => {
        r#"using System; class S{[Obsolete("old")] public string A()=>"a"; public string B()=>A()+"b";} Console.WriteLine(new S().B());"#,
        ["ab"]
    };

    attribute_obsolete_on_interface_implementation => {
        r#"using System; interface I{[Obsolete("old")] string Run();} class S:I{public string Run()=>"ok";} Console.WriteLine(new S().Run());"#,
        ["ok"]
    };

    attribute_flags_or_two_bits => {
        r#"using System; [Flags] enum P{None=0,Read=1,Write=2} Console.WriteLine((int)(P.Read|P.Write));"#,
        ["3"]
    };

    attribute_flags_hasflag_single => {
        r#"using System; [Flags] enum P{Read=1,Write=2} var v=P.Read|P.Write; Console.WriteLine(v.HasFlag(P.Read));"#,
        ["True"]
    };

    attribute_flags_hasflag_missing => {
        r#"using System; [Flags] enum P{Read=1,Write=2,Exec=4} var v=P.Read; Console.WriteLine(v.HasFlag(P.Exec));"#,
        ["False"]
    };

    attribute_flags_to_string_combined => {
        r#"using System; [Flags] enum P{Read=1,Write=2} Console.WriteLine((P.Read|P.Write).ToString());"#,
        ["Read, Write"]
    };

    attribute_flags_none_is_zero => {
        r#"using System; [Flags] enum P{None=0,Read=1} Console.WriteLine((int)P.None);"#,
        ["0"]
    };

    attribute_flags_all_combined => {
        r#"using System; [Flags] enum P{A=1,B=2,C=4} Console.WriteLine((int)(P.A|P.B|P.C));"#,
        ["7"]
    };

    attribute_flags_and_mask => {
        r#"using System; [Flags] enum P{A=1,B=2,C=4} var v=P.A|P.B|P.C; Console.WriteLine((int)(v&P.B));"#,
        ["2"]
    };

    attribute_flags_xor_toggle => {
        r#"using System; [Flags] enum P{A=1,B=2} var v=P.A|P.B; Console.WriteLine((int)(v^P.A));"#,
        ["2"]
    };

    attribute_flags_complement_style => {
        r#"using System; [Flags] enum P{A=1,B=2,C=4} var v=P.A|P.C; Console.WriteLine(v.HasFlag(P.B));"#,
        ["False"]
    };

    attribute_flags_enum_increment_underlying => {
        r#"using System; [Flags] enum P{Read=1,Write=2,Exec=4} Console.WriteLine(P.Exec>P.Read);"#,
        ["True"]
    };

    attribute_serializable_class_is_defined => {
        r#"using System; [Serializable] class Packet{} Console.WriteLine(Attribute.IsDefined(typeof(Packet),typeof(SerializableAttribute)));"#,
        ["True"]
    };

    attribute_serializable_struct_is_defined => {
        r#"using System; [Serializable] struct Point{} Console.WriteLine(Attribute.IsDefined(typeof(Point),typeof(SerializableAttribute)));"#,
        ["True"]
    };

    attribute_serializable_nested_type => {
        r#"using System; class Outer{[Serializable] public class Inner{}} Console.WriteLine(Attribute.IsDefined(typeof(Outer.Inner),typeof(SerializableAttribute)));"#,
        ["True"]
    };

    attribute_serializable_plain_type_false => {
        r#"using System; class Plain{} Console.WriteLine(Attribute.IsDefined(typeof(Plain),typeof(SerializableAttribute)));"#,
        ["False"]
    };

    attribute_serializable_does_not_block_instantiation => {
        r#"using System; [Serializable] class Node{public int Id=3;} Console.WriteLine(new Node().Id);"#,
        ["3"]
    };

    attribute_serializable_with_field_access => {
        r#"using System; [Serializable] class Data{public string Tag="x";} Console.WriteLine(new Data().Tag);"#,
        ["x"]
    };

    attribute_conditional_debug_method_structural => {
        r#"using System; using System.Diagnostics; class Log{[Conditional("DEBUG")] public static void Trace(string m){Console.WriteLine(m);} public static void Run(){Trace("skip"); Console.WriteLine("seen");}} Log.Run();"#,
        ["seen"]
    };

    attribute_conditional_trace_method_structural => {
        r#"using System; using System.Diagnostics; class Log{[Conditional("TRACE")] public static void Mark(){Console.WriteLine("mark");} public static void Run(){Mark(); Console.WriteLine("after");}} Log.Run();"#,
        ["after"]
    };

    attribute_conditional_on_instance_method => {
        r#"using System; using System.Diagnostics; class Log{[Conditional("DEBUG")] public void Trace(){Console.WriteLine("t");} public void Run(){Trace(); Console.WriteLine("r");}} new Log().Run();"#,
        ["r"]
    };

    attribute_conditional_does_not_affect_other_method => {
        r#"using System; using System.Diagnostics; class Log{[Conditional("DEBUG")] public static void A(){} public static void B(){Console.WriteLine("b");} public static void Run(){A(); B();}} Log.Run();"#,
        ["b"]
    };

    attribute_conditional_with_returning_method_structural => {
        r#"using System; using System.Diagnostics; class Calc{[Conditional("DEBUG")] static void Log(int x){Console.WriteLine(x);} public static int Add(int a,int b){Log(a+b); return a+b;}} Console.WriteLine(Calc.Add(2,3));"#,
        ["5"]
    };

    attribute_if_defined_symbol_prints_on_branch => {
        r#"#define VYBETEST_ON
#if VYBETEST_ON
Console.WriteLine("on");
#else
Console.WriteLine("off");
#endif"#,
        ["on"]
    };

    attribute_if_undefined_symbol_prints_off_branch => {
        r#"#if VYBETEST_OFF
Console.WriteLine("on");
#else
Console.WriteLine("off");
#endif"#,
        ["off"]
    };

    attribute_obsolete_and_flags_on_different_types => {
        r#"using System; [Flags] enum F{A=1} [Obsolete("old")] class S{public int Use()=>(int)F.A;} Console.WriteLine(new S().Use());"#,
        ["1"]
    };

    attribute_flags_on_nested_enum => {
        r#"using System; class Host{[Flags] public enum M{A=1,B=2}} Console.WriteLine((int)(Host.M.A|Host.M.B));"#,
        ["3"]
    };

    attribute_serializable_with_obsolete_method => {
        r#"using System; [Serializable] class S{[Obsolete("old")] public string Run()=>"ok";} Console.WriteLine(new S().Run());"#,
        ["ok"]
    };

    attribute_multiple_custom_markers_via_print => {
        r#"using System; [AttributeUsage(AttributeTargets.Class)] class TagAttribute:Attribute{public string Name; public TagAttribute(string n){Name=n;}} [Tag("svc")] class Worker{} var a=(TagAttribute)Attribute.GetCustomAttribute(typeof(Worker),typeof(TagAttribute)); Console.WriteLine(a.Name);"#,
        ["svc"]
    };

    attribute_flags_enum_in_switch_print => {
        r#"using System; [Flags] enum P{Read=1,Write=2} P v=P.Read; string s=v.HasFlag(P.Write)?"w":"r"; Console.WriteLine(s);"#,
        ["r"]
    };

    attribute_obsolete_field_still_readable => {
        r#"using System; class S{[Obsolete("old")] public int N=6;} Console.WriteLine(new S().N);"#,
        ["6"]
    };

    attribute_obsolete_event_subscription_works => {
        r#"using System; class Btn{[Obsolete("old")] public event Action Click; public void Fire(){Click?.Invoke();}} int n=0; var b=new Btn(); b.Click+=()=>n++; b.Fire(); Console.WriteLine(n);"#,
        ["1"]
    };

    attribute_conditional_chained_with_normal_print => {
        r#"using System; using System.Diagnostics; class P{[Conditional("DEBUG")] static void D(){Console.WriteLine("d");} static void N(){Console.WriteLine("n");} public static void Go(){D(); N();}} P.Go();"#,
        ["n"]
    };

    attribute_serializable_inheritance_check => {
        r#"using System; [Serializable] class Base{} class Derived:Base{} Console.WriteLine(Attribute.IsDefined(typeof(Derived),typeof(SerializableAttribute)));"#,
        ["False"]
    };

    attribute_flags_combined_has_all_bits => {
        r#"using System; [Flags] enum P{A=1,B=2,C=4} var v=P.A|P.B|P.C; Console.WriteLine(v.HasFlag(P.A)&&v.HasFlag(P.B)&&v.HasFlag(P.C));"#,
        ["True"]
    };

    attribute_obsolete_constructor_still_runs => {
        r#"using System; class S{[Obsolete("old")] public S(){Console.WriteLine("ctor");}} new S();"#,
        ["ctor"]
    };

    attribute_if_else_nested_structural => {
        r#"#define VYBETEST_A
#if VYBETEST_A
Console.WriteLine("a");
#else
Console.WriteLine("b");
#endif
Console.WriteLine("c");"#,
        ["a", "c"]
    };

    attribute_conditional_on_private_method_structural => {
        r#"using System; using System.Diagnostics; class S{[Conditional("DEBUG")] void Trace(){} public void Run(){Trace(); Console.WriteLine("ok");}} new S().Run();"#,
        ["ok"]
    };

    attribute_flags_underlying_type_cast => {
        r#"using System; [Flags] enum P:byte{Read=1,Write=2} Console.WriteLine((byte)(P.Read|P.Write));"#,
        ["3"]
    };

    attribute_serializable_array_of_instances => {
        r#"using System; [Serializable] class Node{public int Id;} var arr=new Node[]{new Node{Id=1},new Node{Id=2}}; Console.WriteLine(arr[1].Id);"#,
        ["2"]
    };

    attribute_obsolete_indexer_still_works => {
        r#"using System; class S{[Obsolete("old")] public int this[int i]=>i*2;} Console.WriteLine(new S()[4]);"#,
        ["8"]
    };

    attribute_flags_none_hasflag_false => {
        r#"using System; [Flags] enum P{None=0,Read=1} Console.WriteLine(P.None.HasFlag(P.Read));"#,
        ["False"]
    };

    attribute_flags_shift_left_style_combine => {
        r#"using System; [Flags] enum P{A=1,B=2,C=4} var v=(P)0; v|=P.A; v|=P.C; Console.WriteLine((int)v);"#,
        ["5"]
    };

    attribute_combined_if_and_conditional_print => {
        r#"#define VYBETEST_PRE
using System; using System.Diagnostics; class App{[Conditional("DEBUG")] static void Log(){} static void Main(){#if VYBETEST_PRE Console.WriteLine("pre"); #endif Log(); Console.WriteLine("post");}} App.Main();"#,
        ["pre", "post"]
    };
}
