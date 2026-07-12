//! `System.Text.Json`: serialization, deserialization, options.
use super::helpers::run_csharp;

#[test]
fn serialize_simple_object_to_json_string() {
    assert_eq!(
        run_csharp(
            r#"var obj=new{Name="Alice",Age=30};
string json=System.Text.Json.JsonSerializer.Serialize(obj);
Console.WriteLine(json.Contains("Alice"));"#
        ),
        &["True"]
    );
}

#[test]
fn deserialize_json_string_to_typed_record() {
    assert_eq!(
        run_csharp(
            r#"record Person(string Name,int Age);
string json="{\"Name\":\"Bob\",\"Age\":25}";
var p=System.Text.Json.JsonSerializer.Deserialize<Person>(json);
Console.WriteLine(p.Name); Console.WriteLine(p.Age);"#
        ),
        &["Bob", "25"]
    );
}

#[test]
fn json_serialize_list_produces_array_syntax() {
    assert_eq!(
        run_csharp(
            r#"var list=new System.Collections.Generic.List<int>{1,2,3};
string json=System.Text.Json.JsonSerializer.Serialize(list);
Console.WriteLine(json);"#
        ),
        &["[1,2,3]"]
    );
}

#[test]
fn json_options_case_insensitive_deserialization() {
    assert_eq!(
        run_csharp(
            r#"class Item{public string Label{get;set;}}
var opts=new System.Text.Json.JsonSerializerOptions{PropertyNameCaseInsensitive=true};
var item=System.Text.Json.JsonSerializer.Deserialize<Item>("{\"label\":\"x\"}",opts);
Console.WriteLine(item.Label);"#
        ),
        &["x"]
    );
}

#[test]
fn json_deserialize_dictionary_from_object_json() {
    assert_eq!(
        run_csharp(
            r#"var d=System.Text.Json.JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string,int>>("{\"a\":1,\"b\":2}");
Console.WriteLine(d["a"]); Console.WriteLine(d.Count);"#
        ),
        &["1", "2"]
    );
}

#[test]
fn json_roundtrip_preserves_nested_object() {
    assert_eq!(
        run_csharp(
            r#"class Inner{public int X{get;set;}}
class Outer{public Inner Child{get;set;}}
var orig=new Outer{Child=new Inner{X=42}};
var json=System.Text.Json.JsonSerializer.Serialize(orig);
var back=System.Text.Json.JsonSerializer.Deserialize<Outer>(json);
Console.WriteLine(back.Child.X);"#
        ),
        &["42"]
    );
}
