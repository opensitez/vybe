use crate::helpers::*;

#[test]
fn type_alias_shadowing() {
    let out = run_prints(
        "package main; import \"fmt\"; type int string; func main() { var x int = \"hello\"; fmt.Println(x); }",
    );
    assert_eq!(out, vec!["hello"]);
}
#[test]
fn unnamed_struct_compile() {
    compile_ok("package main; func main() { s := struct{ Name string }{Name: \"test\"}; _ = s }");
}
#[test]
fn unnamed_struct_access() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { s := struct{ X int }{X: 5}; fmt.Println(s.X); }",
    );
    assert_eq!(out, vec!["5"]);
}
#[test]
fn type_cast_custom_to_int() {
    let out = run_prints(
        "package main; import \"fmt\"; type MyInt int; func main() { var x MyInt = 10; y := int(x); fmt.Println(y); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn type_cast_int_to_custom() {
    let out = run_prints(
        "package main; import \"fmt\"; type MyInt int; func main() { x := 10; y := MyInt(x); fmt.Println(y); }",
    );
    assert_eq!(out, vec!["10"]);
}
#[test]
fn nested_type_declarations() {
    compile_ok("package main; type A struct { B struct { C int } }; func main() {}");
}
#[test]
fn type_alias_func() {
    compile_ok("package main; type Callback func(int) string; func main() {}");
}
#[test]
fn type_alias_map() {
    compile_ok("package main; type Dict map[string]int; func main() {}");
}
#[test]
fn type_alias_slice() {
    compile_ok("package main; type List []int; func main() {}");
}
#[test]
fn type_alias_pointer() {
    compile_ok("package main; type IntPtr *int; func main() {}");
}
#[test]
fn empty_interface_type_alias() {
    compile_ok("package main; type Any interface{}; func main() {}");
}
#[test]
fn multiple_type_declarations() {
    compile_ok("package main; type ( A int; B string ); func main() {}");
}
#[test]
fn map_of_interface() {
    compile_ok(
        "package main; func main() { m := map[string]interface{}{\"a\": 1, \"b\": \"str\"}; _ = m }",
    );
}
#[test]
fn slice_of_interface() {
    compile_ok("package main; func main() { s := []interface{}{1, \"str\", true}; _ = s }");
}
#[test]
fn type_cast_float_to_uint() {
    let out = run_prints(
        "package main; import \"fmt\"; func main() { f := 3.14; u := uint(f); fmt.Println(u); }",
    );
    assert_eq!(out, vec!["3"]);
}
