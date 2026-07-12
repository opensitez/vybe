//! `using` aliases for types, generic aliases (C# 12), and `global using`.
use super::helpers::run_csharp;

#[test]
fn using_alias_creates_shorter_name_for_type() {
    assert_eq!(
        run_csharp(
            r#"using IntList=System.Collections.Generic.List<int>;
var list=new IntList{1,2,3};
Console.WriteLine(list.Count);"#
        ),
        &["3"]
    );
}

#[test]
fn using_alias_for_fully_qualified_type() {
    assert_eq!(
        run_csharp(
            r#"using Dict=System.Collections.Generic.Dictionary<string,int>;
var d=new Dict{{"a",1},{"b",2}};
Console.WriteLine(d["b"]);"#
        ),
        &["2"]
    );
}

#[test]
fn type_alias_works_as_return_type_and_parameter() {
    assert_eq!(
        run_csharp(
            r#"using NameMap=System.Collections.Generic.Dictionary<string,string>;
NameMap Build()=>new NameMap{{"k","v"}};
Console.WriteLine(Build()["k"]);"#
        ),
        &["v"]
    );
}
