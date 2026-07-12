//! More generic variance: covariant out, contravariant in, wildcard patterns.
use super::helpers::run_csharp;

#[test]
fn covariant_interface_allows_derived_where_base_expected() {
    assert_eq!(
        run_csharp(
            r#"interface IReader<out T>{T Read();}
class StringReader:IReader<string>{public string Read()=>"hello";}
IReader<object> r=new StringReader();
Console.WriteLine(r.Read());"#
        ),
        &["hello"]
    );
}

#[test]
fn contravariant_interface_allows_base_where_derived_expected() {
    assert_eq!(
        run_csharp(
            r#"interface IWriter<in T>{void Write(T v);}
class ObjectWriter:IWriter<object>{public void Write(object v)=>Console.WriteLine(v);}
IWriter<string> w=new ObjectWriter();
w.Write("hi");"#
        ),
        &["hi"]
    );
}

#[test]
fn ienumerable_is_covariant_over_its_element_type() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<string> strings=new[]{"a","b"};
System.Collections.Generic.IEnumerable<object> objects=strings;
Console.WriteLine(objects.Count());"#
        ),
        &["2"]
    );
}

#[test]
fn func_return_type_is_covariant() {
    assert_eq!(
        run_csharp(
            r#"System.Func<string> getStr=()=>"hello";
System.Func<object> getObj=getStr;
Console.WriteLine(getObj());"#
        ),
        &["hello"]
    );
}
