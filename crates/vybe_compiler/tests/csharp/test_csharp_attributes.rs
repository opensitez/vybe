//! Custom attributes: declaration, application, reflection-based retrieval.
use super::helpers::run_csharp;

#[test]
fn custom_attribute_readable_via_get_custom_attributes() {
    assert_eq!(
        run_csharp(r#"[System.AttributeUsage(System.AttributeTargets.Class)]
class TagAttribute:System.Attribute{public string Value;public TagAttribute(string v){Value=v;}}
[Tag("hello")]
class Target{}
var attrs=(TagAttribute[])typeof(Target).GetCustomAttributes(typeof(TagAttribute),false);
Console.WriteLine(attrs[0].Value);"#),
        &["hello"]
    );
}

#[test]
fn attribute_with_named_property_retrieved_correctly() {
    assert_eq!(
        run_csharp(r#"[System.AttributeUsage(System.AttributeTargets.Method)]
class PriorityAttribute:System.Attribute{public int Level{get;set;}}
class Work{
    [Priority(Level=3)]
    public void DoIt(){}
}
var mi=typeof(Work).GetMethod("DoIt");
var attr=(PriorityAttribute)mi.GetCustomAttributes(typeof(PriorityAttribute),false)[0];
Console.WriteLine(attr.Level);"#),
        &["3"]
    );
}

#[test]
fn obsolete_attribute_is_standard_bcl_attribute() {
    assert_eq!(
        run_csharp(r#"class Old{
    [System.Obsolete("use NewMethod")]
    public void OldMethod(){}
}
var mi=typeof(Old).GetMethod("OldMethod");
bool hasObs=mi.GetCustomAttributes(typeof(System.ObsoleteAttribute),false).Length>0;
Console.WriteLine(hasObs);"#),
        &["True"]
    );
}

#[test]
fn attribute_targets_property_applies_to_property() {
    assert_eq!(
        run_csharp(r#"[System.AttributeUsage(System.AttributeTargets.Property)]
class RequiredAttribute:System.Attribute{}
class Form{[Required] public string Name{get;set;}}
var pi=typeof(Form).GetProperty("Name");
bool has=pi.GetCustomAttributes(typeof(RequiredAttribute),false).Length>0;
Console.WriteLine(has);"#),
        &["True"]
    );
}

#[test]
fn multiple_attributes_on_same_target_both_retrieved() {
    assert_eq!(
        run_csharp(r#"[System.AttributeUsage(System.AttributeTargets.Class,AllowMultiple=true)]
class TagAttribute:System.Attribute{public string Name;public TagAttribute(string n){Name=n;}}
[Tag("a")][Tag("b")]
class Thing{}
var attrs=(TagAttribute[])typeof(Thing).GetCustomAttributes(typeof(TagAttribute),false);
Console.WriteLine(attrs.Length);"#),
        &["2"]
    );
}
